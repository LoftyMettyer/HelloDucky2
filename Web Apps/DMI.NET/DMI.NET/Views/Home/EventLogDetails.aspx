<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>

<html>
<head runat="server">

<link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css"/>
<script src="<%: Url.Content("~/Scripts/jquery-1.8.2.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/openhr.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>
    
	<script src="<%: Url.Content("~/Scripts/jquery-ui-1.9.1.custom.min.js") %>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/Scripts/jquery.cookie.js") %>" type="text/javascript"></script>	 	
   <script src="<%: Url.Content("~/Scripts/menu.js")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/Scripts/jquery.ui.touch-punch.min.js") %>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/Scripts/jsTree/jquery.jstree.js") %>" type="text/javascript"></script>
	<script id="officebarscript" src="<%: Url.Content("~/Scripts/officebar/jquery.officebar.js") %>" type="text/javascript"></script>	
    <TITLE>OpenHR Intranet</TITLE>


</head>
<body>

    <%Html.RenderPartial("~/Views/Shared/ctl_ASRIntranetPrintFunctions.ascx")%>

    <OBJECT classid="clsid:F9043C85-F6F2-101A-A3C9-08002B2F49FB" 
	id=dialog 
  codebase="cabs/comdlg32.cab#Version=1,0,0,0"
	style="LEFT: 0px; TOP: 0px" 
	VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="847">
	<PARAM NAME="_ExtentY" VALUE="847">
	<PARAM NAME="_Version" VALUE="393216">
	<PARAM NAME="CancelError" VALUE="0">
	<PARAM NAME="Color" VALUE="0">
	<PARAM NAME="Copies" VALUE="1">
	<PARAM NAME="DefaultExt" VALUE="">
	<PARAM NAME="DialogTitle" VALUE="">
	<PARAM NAME="FileName" VALUE="">
	<PARAM NAME="Filter" VALUE="">
	<PARAM NAME="FilterIndex" VALUE="0">
	<PARAM NAME="Flags" VALUE="0">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="FontName" VALUE="">
	<PARAM NAME="FontSize" VALUE="8">
	<PARAM NAME="FontStrikeThru" VALUE="0">
	<PARAM NAME="FontUnderLine" VALUE="0">
	<PARAM NAME="FromPage" VALUE="0">
	<PARAM NAME="HelpCommand" VALUE="0">
	<PARAM NAME="HelpContext" VALUE="0">
	<PARAM NAME="HelpFile" VALUE="">
	<PARAM NAME="HelpKey" VALUE="">
	<PARAM NAME="InitDir" VALUE="">
	<PARAM NAME="Max" VALUE="0">
	<PARAM NAME="Min" VALUE="0">
	<PARAM NAME="MaxFileSize" VALUE="260">
	<PARAM NAME="PrinterDefault" VALUE="1">
	<PARAM NAME="ToPage" VALUE="0">
	<PARAM NAME="Orientation" VALUE="1">
	</OBJECT>

<form id=frmEventDetails name=frmEventDetails>

<%
	dim rsAllBatchJobs
	dim sSQL
	dim i 
	dim sValue

	dim objUtilities
		
    Dim cmdEventBatchJobs
    Dim prmBatchRunID
    Dim prmEventID
    
    objUtilities = Session("UtilitiesObject")
		
	session("eventName") = Request("txtEventName")
	session("eventID") = Request("txtEventID")
	session("cboString") = vbNullString

    If Request("txtEventMode") = "Batch" Then
        Session("eventBatch") = True
        Response.Write("<INPUT type=hidden Name=txtEventBatch ID=txtEventBatch VALUE=1>" & vbCrLf)
    Else
        Session("eventBatch") = False
        Response.Write("<INPUT type=hidden Name=txtEventBatch ID=txtEventBatch VALUE=0>" & vbCrLf)
    End If

    cmdEventBatchJobs = CreateObject("ADODB.Command")
	cmdEventBatchJobs.CommandText = "spASRIntGetEventLogBatchDetails"
	cmdEventBatchJobs.CommandType = 4 ' Stored procedure
    cmdEventBatchJobs.ActiveConnection = Session("databaseConnection")
								
    prmBatchRunID = cmdEventBatchJobs.CreateParameter("BatchRunID", 3, 1) ' 3=integer, 1=input
    cmdEventBatchJobs.Parameters.Append(prmBatchRunID)
    prmBatchRunID.value = CleanNumeric(Request("txtEventBatchRunID"))

    prmEventID = cmdEventBatchJobs.CreateParameter("EventID", 3, 1) ' 3=integer, 1=input
    cmdEventBatchJobs.Parameters.Append(prmEventID)
	prmEventID.value = cleanNumeric(Request("txtEventID"))

    Err.Clear()
    rsAllBatchJobs = cmdEventBatchJobs.Execute
	
	with rsAllBatchJobs
		if not (.EOF and .BOF) then
			i = 0
			do until .EOF
				i = i + 1

                Response.Write("<INPUT type=hidden Name=txtEventID_" & .Fields("ID").Value & " ID=txtEventID_" & .Fields("ID").Value & " VALUE=""" & Replace(.Fields("ID").Value, """", "&quot;") & """>" & vbCrLf)
				
                sValue = .Fields("Name").Value                      'original value
                sValue = Replace(sValue, """", "&quot;")    'escape quotes
                sValue = Replace(sValue, "<", "&lt;")           'escape left angle bracket
                sValue = Replace(sValue, ">", "&gt;")           'escape right angle bracket
				
                Response.Write("<INPUT type=hidden Name=txtEventName_" & .Fields("ID").Value & " ID=txtEventName_" & .Fields("ID").Value & " VALUE=""" & sValue & """>" & vbCrLf)
                Response.Write("<INPUT type=hidden Name=txtEventMode_" & .Fields("ID").Value & " ID=txtEventMode_" & .Fields("ID").Value & " VALUE=""" & Replace(.Fields("Mode").Value, """", "&quot;") & """>" & vbCrLf)
				
                Response.Write("<INPUT type=hidden Name=txtEventStartTime_" & .Fields("ID").Value & " ID=txtEventStartTime_" & .Fields("ID").Value & " VALUE=""" & ConvertSqlDateToLocale(.Fields("DateTime").Value) & " " & ConvertSqlDateToTime(.Fields("DateTime").Value) & """>" & vbCrLf)
				
                If IsDBNull(.Fields("EndTime").Value) Then
                    Response.Write("<INPUT type=hidden Name=txtEventEndTime_" & .Fields("ID").Value & " ID=txtEventEndTime_" & .Fields("ID").Value & " VALUE=""" & vbNullString & """>" & vbCrLf)
                Else
                    Response.Write("<INPUT type=hidden Name=txtEventEndTime_" & .Fields("ID").Value & " ID=txtEventEndTime_" & .Fields("ID").Value & " VALUE=""" & ConvertSqlDateToLocale(.Fields("EndTime").Value) & " " & ConvertSqlDateToTime(.Fields("EndTime").Value) & """>" & vbCrLf)
                End If
				
                Response.Write("<INPUT type=hidden Name=txtEventDuration_" & .Fields("ID").Value & " ID=txtEventDuration_" & .Fields("ID").Value & " VALUE=""" & objUtilities.FormatEventDuration(CLng(.Fields("Duration").Value)) & """>" & vbCrLf)

                Response.Write("<INPUT type=hidden Name=txtEventType_" & .Fields("ID").Value & " ID=txtEventType_" & .Fields("ID").Value & " VALUE=""" & Replace(.Fields("Type").Value, """", "&quot;") & """>" & vbCrLf)
                Response.Write("<INPUT type=hidden Name=txtEventStatus_" & .Fields("ID").Value & " ID=txtEventStatus_" & .Fields("ID").Value & " VALUE=""" & Replace(.Fields("Status").Value, """", "&quot;") & """>" & vbCrLf)
                Response.Write("<INPUT type=hidden Name=txtEventUser_" & .Fields("ID").Value & " ID=txtEventUser_" & .Fields("ID").Value & " VALUE=""" & Replace(.Fields("Username").Value, """", "&quot;") & """>" & vbCrLf)
				
                Response.Write("<INPUT type=hidden Name=txtEventSuccessCount_" & .Fields("ID").Value & " ID=txtEventSuccessCount_" & .Fields("ID").Value & " VALUE=""" & Replace(.Fields("SuccessCount").Value, """", "&quot;") & """>" & vbCrLf)
                Response.Write("<INPUT type=hidden Name=txtEventFailCount_" & .Fields("ID").Value & " ID=txtEventFailCount_" & .Fields("ID").Value & " VALUE=""" & Replace(.Fields("FailCount").Value, """", "&quot;") & """>" & vbCrLf)
				
                Response.Write("<INPUT type=hidden Name=txtEventBatchRunID_" & .Fields("ID").Value & " ID=txtEventBatchRunID_" & .Fields("ID").Value & " VALUE=""" & Replace(.Fields("BatchRunID").Value, """", "&quot;") & """>" & vbCrLf)
                Response.Write("<INPUT type=hidden Name=txtEventBatchName_" & .Fields("ID").Value & " ID=txtEventBatchName_" & .Fields("ID").Value & " VALUE=""" & Replace(.Fields("BatchName").Value, """", "&quot;") & """>" & vbCrLf)
                Response.Write("<INPUT type=hidden Name=txtEventBatchJobID_" & .Fields("ID").Value & " ID=txtEventBatchJobID_" & .Fields("ID").Value & " VALUE=""" & Replace(.Fields("BatchJobID").Value, """", "&quot;") & """>" & vbCrLf)
				
                If Session("eventBatch") = True Then
                    If Session("eventID") = .Fields("ID").Value Then
                        Session("cboString") = Session("cboString") & "<OPTION SELECTED NAME='" & .Fields("Name").Value & "' VALUE='" & .Fields("ID").Value & "'>" & .Fields("Type").Value & " - " & .Fields("Name").Value & vbCrLf
                    Else
                        Session("cboString") = Session("cboString") & "<OPTION NAME='" & .Fields("Name").Value & "' VALUE='" & .Fields("ID").Value & "'>" & .Fields("Type").Value & " - " & .Fields("Name").Value & vbCrLf
                    End If
                End If
				
                .MoveNext()
            Loop
			
            Session("eventBatchName") = Request("txtEventBatchName")
			
            Session("cboString") = Session("cboString") & "</SELECT>" & vbCrLf
            If i <= 1 Then
                Session("cboString") = "<select disabled id=cboOtherJobs name=cboOtherJobs class=""combodisabled"" style=""WIDTH: 100%"" onchange='populateEventInfo();populateEventDetails();'>" & vbCrLf & Session("cboString")
            Else
                Session("cboString") = "<select id=cboOtherJobs name=cboOtherJobs class=""combo"" style=""WIDTH: 100%"" onchange='populateEventInfo();populateEventDetails();'>" & vbCrLf & Session("cboString")
            End If
        End If
    End With
	
    rsAllBatchJobs = Nothing
    cmdEventBatchJobs = Nothing
    prmEventID = Nothing
    prmBatchRunID = Nothing
    objUtilities = Nothing
	
    Response.Write("<INPUT type=hidden Name=txtOriginalEventID ID=txtOriginalEventID VALUE=" & Request("txtEventID") & ">" & vbCrLf)
%>
<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD >
			<TABLE WIDTH="100%" height="100%" class="invisible" cellspacing=0 cellpadding=0>
				<tr height=5> 
					<td colspan=3></td>
				</tr> 
								
				<tr> 
					<TD width=5></td>
					<td>
						<TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
							<tr  valign=top> 
								<td>
									<TABLE HEIGHT="100%" WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
<!--										<TR height=5>
											<TD>
												Details : 
											</TD>
										</TR>
-->					
										<TR>
											<TD>
												<TABLE HEIGHT="100%" WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=4>
													
<%					

	if session("eventBatch") = true	then					
        Response.Write("										<TR height=20> " & vbCrLf)
        Response.Write("												<td>" & vbCrLf)
        Response.Write("													<TABLE WIDTH='100%' class=""invisible"" CELLSPACING=0 CELLPADDING=4>" & vbCrLf)
        Response.Write("														<TR> " & vbCrLf)
        Response.Write("															<TD width=120 nowrap>" & vbCrLf)
        Response.Write("																Batch Job Name :  " & vbCrLf)
        Response.Write("															</TD> " & vbCrLf)
        Response.Write("															<TD width=200 NAME=tdBatchJobName ID=tdBatchJobName> " & vbCrLf)
        Response.Write("																" & Session("eventBatchName") & vbCrLf)
        Response.Write("															</TD>" & vbCrLf)
        Response.Write("															<TD width=120 nowrap> " & vbCrLf)
        Response.Write("																All Jobs in Batch :  " & vbCrLf)
        Response.Write("															</TD>" & vbCrLf)
        Response.Write("															<TD> " & vbCrLf)
		
        Response.Write(Session("cboString"))
		
        Response.Write("															</TD> " & vbCrLf)
        Response.Write("														</TR>" & vbCrLf)
        Response.Write("													</TABLE>" & vbCrLf)
        Response.Write("												</TD>" & vbCrLf)
        Response.Write("											</TR>" & vbCrLf)
        Response.Write("											<TR height=10> " & vbCrLf)
        Response.Write("												<TD>" & vbCrLf)
        Response.Write("													<hr> " & vbCrLf)
        Response.Write("												</TD>" & vbCrLf)
        Response.Write("											</TR>" & vbCrLf)
    End If
%>												
													<TR height=10>
														<td>
															<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=4>
																<TR>
																	<TD nowrap>
																		Name : 
																	</TD>
																	<TD NAME=tdame ID=tdName colspan=3>
																		
																	</TD>
																	
																	<TD width="8%" nowrap>
																		Mode : 
																	</TD>
																	<TD width="35%" NAME=tdMode ID=tdMode>
																		
																	</TD>
																</TR>
																
																<TR height=5>
																	<TD colspan=6></TD>
																</TR>
																
																<TR height=10>
																	<TD width="8%" nowrap>
																		Start : 
																	</TD>
																	<TD width="25%" NAME=tdStartTime ID=tdStartTime>
																		
																	</TD>
																	<TD width="8%" nowrap>
																		End : 
																	</TD>
																	<TD width="25%" NAME=tdEndTime ID=tdEndTime>
																		
																	</TD>
																	<TD width="8%" nowrap>
																		Duration : 
																	</TD>
																	<TD width="25%" NAME=tdDuration ID=tdDuration>
																		
																	</TD>
																</TR>
																
																<TR height=5>
																	<TD colspan=6></TD>
																</TR>
													
																<TR height=10>
																	<TD nowrap>
																		Type : 
																	</TD>
																	<TD NAME=tdType ID=tdType>
																		
																	</TD>
																	<TD nowrap>
																		Status : 
																	</TD>
																	<TD NAME=tdStatus ID=tdStatus>

																	</TD>
																	<TD width="9%" nowrap>
																		User name : 
																	</TD>
																	<TD width="15%" NAME=tdUser ID=tdUser>
																		
																	</TD>
																</TR>
															</TABLE>
														</td>
													</tr>
													<TR height=10>
														<TD>
															<hr>
														</TD>
													</TR>
													<TR height=10>
														<td>
															<TABLE WIDTH='100%' class="invisible" CELLSPACING=0 CELLPADDING=4> 
																<TR>
																	<TD width="23%" nowrap>
																		Records Successful : 
																	</TD>
																	<TD width="10%" NAME=tdSuccessCount ID=tdSuccessCount>
																		
																	</TD>
																	<TD width="20%" nowrap>
																		Records Failed : 
																	</TD>
																	<TD width="50%" NAME=tdFailCount ID=tdFailCount>
																		
																	</TD>
																</TR>
															</TABLE>
														</td>
													</TR>
													<TR height=5>
														<TD colspan=6></TD>
													</TR>
													<TR>
														<TD colspan=6 ID=gridCell Name=gridCell>
															<OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
																	  codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" 
																	height="100%" 
																	id=ssOleDBGridEventLogDetails 
																	name=ssOleDBGridEventLogDetails
																	style="HEIGHT: 100%; VISIBILITY: visible; WIDTH: 100%" 
																	width="100%">
																<PARAM NAME="ScrollBars" VALUE="2">
																<PARAM NAME="_Version" VALUE="196617">
																<PARAM NAME="DataMode" VALUE="2">
																<PARAM NAME="Cols" VALUE="0">
																<PARAM NAME="Rows" VALUE="0">
																<PARAM NAME="BorderStyle" VALUE="1">
																<PARAM NAME="RecordSelectors" VALUE="0">
																<PARAM NAME="GroupHeaders" VALUE="-1">
																<PARAM NAME="ColumnHeaders" VALUE="-1">
																<PARAM NAME="GroupHeadLines" VALUE="1">
																<PARAM NAME="HeadLines" VALUE="1">
																<PARAM NAME="FieldDelimiter" VALUE="(None)">
																<PARAM NAME="FieldSeparator" VALUE="(Tab)">
																<PARAM NAME="Row.Count" VALUE="0">
																<PARAM NAME="Col.Count" VALUE="1">
																<PARAM NAME="stylesets.count" VALUE="0">
																<PARAM NAME="TagVariant" VALUE="EMPTY">
																<PARAM NAME="UseGroups" VALUE="0">
																<PARAM NAME="HeadFont3D" VALUE="0">
																<PARAM NAME="Font3D" VALUE="0">
																<PARAM NAME="DividerType" VALUE="3">
																<PARAM NAME="DividerStyle" VALUE="1">
																<PARAM NAME="DefColWidth" VALUE="0">
																<PARAM NAME="BeveColorScheme" VALUE="2">
																<PARAM NAME="BevelColorFrame" VALUE="0">
																<PARAM NAME="BevelColorHighlight" VALUE="0">
																<PARAM NAME="BevelColorShadow" VALUE="0">
																<PARAM NAME="BevelColorFace" VALUE="0">
																<PARAM NAME="CheckBox3D" VALUE="-1">
																<PARAM NAME="AllowAddNew" VALUE="0">
																<PARAM NAME="AllowDelete" VALUE="0">
																<PARAM NAME="AllowUpdate" VALUE="0">
																<PARAM NAME="MultiLine" VALUE="-1">
																<PARAM NAME="ActiveCellStyleSet" VALUE="">
																<PARAM NAME="RowSelectionStyle" VALUE="0">
																<PARAM NAME="AllowRowSizing" VALUE="-1">
																<PARAM NAME="AllowGroupSizing" VALUE="0">
																<PARAM NAME="AllowColumnSizing" VALUE="-1">
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
																<PARAM NAME="CellNavigation" VALUE="1">
																<PARAM NAME="MaxSelectedRows" VALUE="1">
																<PARAM NAME="HeadStyleSet" VALUE="">
																<PARAM NAME="StyleSet" VALUE="">
																<PARAM NAME="ForeColorEven" VALUE="0">
																<PARAM NAME="ForeColorOdd" VALUE="0">
																<PARAM NAME="BackColorEven" VALUE="0">
																<PARAM NAME="BackColorOdd" VALUE="0">
																<PARAM NAME="Levels" VALUE="1">
																<PARAM NAME="RowHeight" VALUE="1000">
																<PARAM NAME="ExtraHeight" VALUE="0">
																<PARAM NAME="ActiveRowStyleSet" VALUE="">
																<PARAM NAME="CaptionAlignment" VALUE="2">
																<PARAM NAME="SplitterPos" VALUE="0">
																<PARAM NAME="SplitterVisible" VALUE="0">
																<PARAM NAME="Columns.Count" VALUE="1">
																<!--Details-->        
																<PARAM NAME="Columns(0).Width" VALUE="17000">
																<PARAM NAME="Columns(0).Visible" VALUE="-1">
																<PARAM NAME="Columns(0).Columns.Count" VALUE="1">
																<PARAM NAME="Columns(0).Caption" VALUE="Details">
																<PARAM NAME="Columns(0).Name" VALUE="Details">
																<PARAM NAME="Columns(0).Alignment" VALUE="0">
																<PARAM NAME="Columns(0).CaptionAlignment" VALUE="2">
																<PARAM NAME="Columns(0).Bound" VALUE="0">
																<PARAM NAME="Columns(0).AllowSizing" VALUE="1">
																<PARAM NAME="Columns(0).DataField" VALUE="Column 0">
																<PARAM NAME="Columns(0).DataType" VALUE="8">
																<PARAM NAME="Columns(0).Level" VALUE="0">
																<PARAM NAME="Columns(0).NumberFormat" VALUE="">
																<PARAM NAME="Columns(0).Case" VALUE="0">
																<PARAM NAME="Columns(0).FieldLen" VALUE="256">
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
																<PARAM NAME="_ExtentX" VALUE="11298">
																<PARAM NAME="_ExtentY" VALUE="3969">
																<PARAM NAME="_StockProps" VALUE="79">
																<PARAM NAME="Caption" VALUE="">
																<PARAM NAME="ForeColor" VALUE="0">
																<PARAM NAME="BackColor" VALUE="0">
																<PARAM NAME="Enabled" VALUE="-1">
																<PARAM NAME="DataMember" VALUE="">
															</OBJECT>
														</TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<tr height=5>
											<td align=right valign=bottom>
												<TABLE class="invisible" CELLSPACING=0 CELLPADDING=4>
													<TR>
														<TD width=10>
															<INPUT id=cmdEmail type=button class="btn" value="Email..." name=cmdEmail style="WIDTH: 80px" width="80"
															    onclick="emailEvent();" 
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
														<TD width=5>
															<INPUT id=cmdPrint class="btn" type=button value="Print..." name=cmdPrint style="WIDTH: 80px" width="80"
															    onclick="printEvent(true);" 
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
														<TD width=5>
															<INPUT id=cmdOK type=button class="btn" value=OK name=cmdOK style="WIDTH: 80px" width="80" 
															    onclick="okClick();"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
													</tr>
												</table>								
											</td>
										</tr>		
									</TABLE>
								</td>
							</tr>
						</TABLE>
					</td>
					<TD width=5></td>
				</tr> 
			</TABLE>
		</td>
	</tr> 
</TABLE>
</form>

<form id=frmDetails name=frmDetails style="visibility:hidden;display:none">
<%
	dim iDetailCount
    Dim rsEventDetails
    Dim cmdEventDetails
    Dim prmEventExists    

	iDetailCount = 0
	
    cmdEventDetails = CreateObject("ADODB.Command")
	cmdEventDetails.CommandText = "spASRIntGetEventLogDetails"
	cmdEventDetails.CommandType = 4 ' Stored procedure
    cmdEventDetails.ActiveConnection = Session("databaseConnection")
								
    prmBatchRunID = cmdEventDetails.CreateParameter("BatchRunID", 3, 1) ' 3=integer, 1=input
    cmdEventDetails.Parameters.Append(prmBatchRunID)
	prmBatchRunID.value = cleanNumeric(Request("txtEventBatchRunID"))

    prmEventID = cmdEventDetails.CreateParameter("EventID", 3, 1) ' 3=integer, 1=input
    cmdEventDetails.Parameters.Append(prmEventID)
	prmEventID.value = cleanNumeric(Request("txtEventID"))

    prmEventExists = cmdEventDetails.CreateParameter("EventExists", 3, 2) ' 3=integer, 2=output
    cmdEventDetails.Parameters.Append(prmEventExists)

    Err.Clear()
    rsEventDetails = cmdEventDetails.Execute

	if not (rsEventDetails.BOF and rsEventDetails.EOF) then
		do while not rsEventDetails.EOF
			iDetailCount = iDetailCount + 1
				
            sValue = rsEventDetails.Fields("Notes").value       'original value
			sValue = Replace(sValue, """", "&quot;")	'escape quotes
			
            Response.Write("<INPUT type=hidden Name=txtEventNotes_" & rsEventDetails.Fields("EventLogID").value & "_" & iDetailCount & " ID=txtEventNotes_" & rsEventDetails.Fields("EventLogID").value & "_" & iDetailCount & " VALUE=""" & sValue & """>" & vbCrLf)
				
			rsEventDetails.MoveNext 
		loop	
	end if
	rsEventDetails.close
    rsEventDetails = Nothing
	
    If cmdEventDetails.Parameters("EventExists").Value > 0 Then
        Response.Write("<INPUT TYPE=hidden NAME=txtEventExists ID=txtEventExists VALUE='1'>" & vbCrLf)
    Else
        Response.Write("<INPUT TYPE=hidden NAME=txtEventExists ID=txtEventExists VALUE='0'>" & vbCrLf)
    End If
	
    cmdEventDetails = Nothing
    prmEventID = Nothing
    prmBatchRunID = Nothing
    prmEventExists = Nothing
%>
</form>

<FORM id=frmUseful name=frmUseful style="visibility:hidden;display:none">
	<INPUT type="hidden" id=txtUserName name=txtUserName value="<%=session("username")%>">
<%
    Dim cmdDefinition
    Dim prmModuleKey
    Dim prmParameterKey
    Dim prmParameterValue
    Dim sErrorDescription
    
    cmdDefinition = CreateObject("ADODB.Command")
	cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
	cmdDefinition.CommandType = 4 ' Stored procedure.
    cmdDefinition.ActiveConnection = Session("databaseConnection")

    prmModuleKey = cmdDefinition.CreateParameter("moduleKey", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDefinition.Parameters.Append(prmModuleKey)
	prmModuleKey.value = "MODULE_PERSONNEL"

    prmParameterKey = cmdDefinition.CreateParameter("paramKey", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDefinition.Parameters.Append(prmParameterKey)
	prmParameterKey.value = "Param_TablePersonnel"

    prmParameterValue = cmdDefinition.CreateParameter("paramValue", 200, 2, 8000) '200=varchar, 2=output, 8000=size
    cmdDefinition.Parameters.Append(prmParameterValue)

    Err.Clear()
	cmdDefinition.Execute

    Response.Write("<INPUT type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").Value & ">" & vbCrLf)
	
    cmdDefinition = Nothing

    Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
%>
</FORM>

<FORM id=frmEmail name=frmEmail method=post style="visibility:hidden;display:none" action="emailSelection.asp">
	<INPUT type="hidden" id=txtSelectedEventIDs name=txtSelectedEventIDs>
	<INPUT type="hidden" id=txtBatchInfo name=txtBatchInfo>
	<INPUT type="hidden" id=txtBatchy name=txtBatchy value=0>
	<INPUT type="hidden" id=txtFromMain name=txtFromMain value=0>
</FORM>

    
    <script type="text/javascript">

        function eventlogdetails_window_onload() {

            setGridFont(frmEventDetails.ssOleDBGridEventLogDetails);

            if (frmDetails.txtEventExists.value == 0) {
                
                var frmOpenerRefresh =  window.dialogArguments.OpenHR.getForm("workframe","frmRefresh");
                var frmMainLog =  window.dialogArguments.OpenHR.getForm("workframe","frmLog");

                OpenHR.messageBox("This record no longer exists in the event log.", 48, "Event Log");

                frmOpenerRefresh.txtCurrentUsername.value = frmMainLog.cboUsername.options[frmMainLog.cboUsername.selectedIndex].value;
                frmOpenerRefresh.txtCurrentType.value = frmMainLog.cboType.options[frmMainLog.cboType.selectedIndex].value;
                frmOpenerRefresh.txtCurrentMode.value = frmMainLog.cboMode.options[frmMainLog.cboMode.selectedIndex].value;
                frmOpenerRefresh.txtCurrentStatus.value = frmMainLog.cboStatus.options[frmMainLog.cboStatus.selectedIndex].value;

                frmOpenerRefresh.submit();

                self.close();
            } else {
                var frmOpenerDetails =  window.dialogArguments.OpenHR.getForm("workframe","frmDetails");

                if (frmOpenerDetails.txtEmailPermission.value == 1) {
                    button_disable(frmEventDetails.cmdEmail, false);
                } else {
                    button_disable(frmEventDetails.cmdEmail, true);
                }

                populateEventInfo();

                populateEventDetails();
            }
        }
    </script>

        <script type="text/javascript" id=scptGeneralFunctions>
<!--

    function okClick()
    {
        self.close();
    }

    function emailEvent()
    {
        var sBatchInfo = "";
        var sURL;
	
        if (frmEventDetails.txtEventBatch.value == 1)
        {
            frmEmail.txtBatchy.value = 1;
            frmEmail.txtSelectedEventIDs.value = frmEventDetails.cboOtherJobs.options[frmEventDetails.cboOtherJobs.selectedIndex].value;
		
            sBatchInfo = sBatchInfo + "Batch Job Name :	" + document.getElementById('tdBatchJobName').innerText + String.fromCharCode(13) + String.fromCharCode(13);
		
            sBatchInfo = sBatchInfo + "All Jobs in Batch :	" + String.fromCharCode(13) + String.fromCharCode(13);	
	
            for (var iCount=0; iCount < frmEventDetails.cboOtherJobs.options.length; iCount++)
            {
                sBatchInfo = sBatchInfo + String(frmEventDetails.cboOtherJobs.options[iCount].text) + String.fromCharCode(13) + String.fromCharCode(13);
            }
        }
        else
        {
            frmEmail.txtBatchy.value = 0;
            frmEmail.txtSelectedEventIDs.value = frmEventDetails.txtOriginalEventID.value;
        }
	
        frmEmail.txtBatchInfo.value = sBatchInfo;
	
        sURL = "emailSelection" +
            "?txtSelectedEventIDs=" + frmEmail.txtSelectedEventIDs.value +
            "&txtEmailOrderColumn=" +
            "&txtEmailOrderOrder=" +            
            "&txtFromMain=" + frmEmail.txtFromMain.value + 
            "&txtBatchInfo=" + escape(frmEmail.txtBatchInfo.value) + 
            "&txtBatchy=" + frmEmail.txtBatchy.value;
        openDialog(sURL, 435, 350);
    }
	
    function printEvent(pfToPrinter)
    {
        var fOK = true;
        var sErrorString = new String("");
        var iCurrentRec;
        var sCR = String.fromCharCode(13);
        var sLF = String.fromCharCode(10);
        var objPrinter = ASRIntranetPrintFunctions;
	
        if (pfToPrinter == true) 
        {
            if(objPrinter.IsOK == false) 
            {
                return;
            }
        }
	
        // OK so far.
        if (pfToPrinter == true) 
        {
            fOK = objPrinter.PrintStart(false, frmUseful.txtUserName.value);
        }
	
        if (fOK == true) 
        {	

            if (pfToPrinter == true) 
            {
                //print selected event information
                objPrinter.PrintHeader("Event Log : " + document.getElementById('tdName').innerText);
                objPrinter.PrintNonBold("Mode :	" + document.getElementById('tdMode').innerText);
                objPrinter.PrintNormal("");
                objPrinter.PrintNonBold("Start Time :	" + document.getElementById('tdStartTime').innerText);
                objPrinter.PrintNonBold("End Time :	" + document.getElementById('tdEndTime').innerText);
                objPrinter.PrintNonBold("End Time :	" + document.getElementById('tdDuration').innerText);
                objPrinter.PrintNormal("");
                objPrinter.PrintNonBold("Type :	" + document.getElementById('tdType').innerText);
                objPrinter.PrintNonBold("Status :	" + document.getElementById('tdStatus').innerText);
                objPrinter.PrintNonBold("User name :	" + document.getElementById('tdUser').innerText);
                objPrinter.PrintNormal("");
            }
		    
            if (pfToPrinter == true && (frmEventDetails.txtEventBatch.value == 1)) 
            {
                //print batch job information
                objPrinter.PrintNonBold("Batch Job Name :	" + document.getElementById('tdBatchJobName').innerText);
                objPrinter.PrintNormal("");
                objPrinter.PrintNormal("All Jobs in Batch :");
                objPrinter.PrintNormal("");
			
                for (var iCount=0; iCount < frmEventDetails.cboOtherJobs.options.length; iCount++)
                {
                    objPrinter.PrintNonBold(frmEventDetails.cboOtherJobs.options[iCount].text);
                }
            }
		
            if (pfToPrinter == true)
            {
                //print records summary information			
                objPrinter.PrintNormal("");
                objPrinter.PrintNonBold("Records Successful :	" + document.getElementById('tdSuccessCount').innerText);
                objPrinter.PrintNonBold("Records Failed :	" + document.getElementById('tdFailCount').innerText);
            }
		
            if (pfToPrinter == true)
            {
                //print selected event details
                objPrinter.PrintNormal("");
                objPrinter.PrintBold("Details : ");
                objPrinter.PrintNormal("");
			
                if (frmEventDetails.ssOleDBGridEventLogDetails.Rows < 1)
                {
                    objPrinter.PrintNonBold("There are no details for this event log entry");
                }
                else
                {
                    frmEventDetails.ssOleDBGridEventLogDetails.Redraw = false;
                    frmEventDetails.ssOleDBGridEventLogDetails.MoveFirst();
                    for (var i=0; i < frmEventDetails.ssOleDBGridEventLogDetails.Rows; i++)
                    {
                        iCurrentRec = i + 1;
                        objPrinter.PrintBold("*** Log entry " + iCurrentRec + " of " + frmEventDetails.ssOleDBGridEventLogDetails.Rows + " ***");
					
                        sErrorString = frmEventDetails.ssOleDBGridEventLogDetails.Columns(0).Text;
					
                        objPrinter.PrintNonBold(sErrorString);
                        objPrinter.PrintNormal("");
										
                        frmEventDetails.ssOleDBGridEventLogDetails.MoveNext();
                    }
                    frmEventDetails.ssOleDBGridEventLogDetails.Redraw = true;
                }
            }
			
            if (pfToPrinter == true) 
            {
                objPrinter.PrintEnd();
                objPrinter.PrintConfirm("Event Log Details", "Event Log Details");
            }
        }
    }
	
    function populateEventInfo()
    {
        var sNumber;
	
        if (frmEventDetails.txtEventBatch.value == true) 
        {
            sNumber = frmEventDetails.cboOtherJobs.options[frmEventDetails.cboOtherJobs.selectedIndex].value;
        }
        else
        {
            sNumber = frmEventDetails.txtOriginalEventID.value;
        }

        document.getElementById('tdName').innerHTML = document.getElementById('txtEventName_' + sNumber).value;
        document.getElementById('tdMode').innerHTML = document.getElementById('txtEventMode_' + sNumber).value;

        document.getElementById('tdStartTime').innerHTML = document.getElementById('txtEventStartTime_' + sNumber).value;
        document.getElementById('tdEndTime').innerHTML = document.getElementById('txtEventEndTime_' + sNumber).value;
        document.getElementById('tdDuration').innerHTML = document.getElementById('txtEventDuration_' + sNumber).value;
	
        //document.getElementById('tdTime').innerHTML = ASRIntranetFunctions.ConvertSQLDateToTime(document.getElementById('txtEventTime_' + sNumber).value);
        document.getElementById('tdType').innerHTML = document.getElementById('txtEventType_' + sNumber).value;
        document.getElementById('tdStatus').innerHTML = document.getElementById('txtEventStatus_' + sNumber).value;
        document.getElementById('tdUser').innerHTML = document.getElementById('txtEventUser_' + sNumber).value;

        document.getElementById('tdSuccessCount').innerHTML = document.getElementById('txtEventSuccessCount_' + sNumber).value;
        document.getElementById('tdFailCount').innerHTML = document.getElementById('txtEventFailCount_' + sNumber).value;
    }

    function populateEventDetails()
    {
        var sNumber;
        var iIndex;
        var sControlName;
        var sControl;
        var sAddLine;

        with (frmEventDetails.ssOleDBGridEventLogDetails)
        {
            if (frmDetails.elements.length > 0)
            {
                focus();
                Redraw = false;
                //Reference 4510 - RemoveAll was causing grid to error. 
                if(Rows > 0)
                {
                    RemoveAll();
                }
				
                for (var i=0; i<frmDetails.elements.length; i++)
                {
                    sControl = frmDetails.elements[i];
                    sControlName = frmDetails.elements[i].name;
				
                    if (sControlName != "txtEventExists")
                    {
                        sNumber = sControlName.substr(sControlName.indexOf("_") + 1, sControlName.length);
                        sNumber = sNumber.substr(0, sNumber.indexOf("_"));

                        if (frmEventDetails.txtEventBatch.value == 1)
                        {
                            if (sNumber == frmEventDetails.cboOtherJobs.options[frmEventDetails.cboOtherJobs.selectedIndex].value)
                            {
                                sAddLine = sControl.value;
                                AddItem(sAddLine);
                            }
                        }
                        else
                        {
                            if (sNumber == frmEventDetails.txtOriginalEventID.value)
                            {
                                sAddLine = sControl.value;
                                AddItem(sAddLine); 
                            }
                        }
                    }
                }
                Redraw = true;
            }
            RowHeight = 100; 
        }

        //setGridCaption();
    }

    function setGridCaption()
    {
        var iCurrRec;
        var iTotalRec;
	
        //Update the grid caption after the user has used keys to view the details
        if (frmEventDetails.ssOleDBGridEventLogDetails.Rows == 0)
        {
            frmEventDetails.ssOleDBGridEventLogDetails.Columns("Details").Caption = "No details exist for this entry";
            frmEventDetails.ssOleDBGridEventLogDetails.Enabled = false;
        }
        else
        {
            frmEventDetails.ssOleDBGridEventLogDetails.Enabled = true;
            iCurrRec = parseInt(frmEventDetails.ssOleDBGridEventLogDetails.AddItemRowIndex(frmEventDetails.ssOleDBGridEventLogDetails.Bookmark)) + 1;
            iTotalRec = frmEventDetails.ssOleDBGridEventLogDetails.Rows;
            frmEventDetails.ssOleDBGridEventLogDetails.Columns("Details").Caption = "Details (" + iCurrRec + " Of " + iTotalRec + " Entries)";
        }
    }

    function openDialog(pDestination, pWidth, pHeight)
    {
        dlgwinprops = "center:yes;" +
            "dialogHeight:" + pHeight + "px;" +
            "dialogWidth:" + pWidth + "px;" +
            "help:no;" +
            "resizable:yes;" +
            "scroll:yes;" +
            "status:no;";
        window.showModalDialog(pDestination, self, dlgwinprops);
    }
	
    -->
</script>


    <script type="text/javascript">
        eventlogdetails_window_onload();
    </script>

</body>


</html>
