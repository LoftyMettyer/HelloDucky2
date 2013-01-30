<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>

<html>
<head runat="server">   
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

    <title>Event Log Selection - OpenHR Intranet</title>
</head>
<body>
    
    <OBJECT 
	classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" 
	id="Microsoft_Licensed_Class_Manager_1_0" 
	VIEWASTEXT>
	<PARAM NAME="LPKPath" VALUE="lpks/main.lpk">
</OBJECT>
      
<form id=frmEmail>
<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD>
			<TABLE WIDTH="100%" height="100%" class="invisible" cellspacing=0 cellpadding=0>
				<tr height=5> 
					<td colspan=3></td>
				</tr> 
				<tr> 
					<TD width=5></td>
					<td>
						<TABLE WIDTH="100%" height="100%" class="invisible" cellspacing=4 cellpadding=0>
							<TR>
								<TD width=100%>
									<OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
											 codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" 
											height="100%" 
											id=ssOleDBGridEmail 
											name=ssOleDBGridEmail 
											style="HEIGHT: 100%; VISIBILITY: visible; WIDTH: 100%" 
											width="100%">
										<PARAM NAME="ScrollBars" VALUE="4">
										<PARAM NAME="_Version" VALUE="196617">
										<PARAM NAME="DataMode" VALUE="2">
										<PARAM NAME="Cols" VALUE="0">
										<PARAM NAME="Rows" VALUE="0">
										<PARAM NAME="BorderStyle" VALUE="1">
										<PARAM NAME="RecordSelectors" VALUE="0">
										<PARAM NAME="GroupHeaders" VALUE="-1">
										<PARAM NAME="ColumnHeaders" VALUE="-1">
										<PARAM NAME="GroupHeadLines" VALUE="1">
										<PARAM NAME="HeadLines" VALUE="2">
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
										<PARAM NAME="CheckBox3D" VALUE="0">
										<PARAM NAME="AllowAddNew" VALUE="0">
										<PARAM NAME="AllowDelete" VALUE="0">
										<PARAM NAME="AllowUpdate" VALUE="-1">
										<PARAM NAME="MultiLine" VALUE="0">
										<PARAM NAME="ActiveCellStyleSet" VALUE="">
										<PARAM NAME="RowSelectionStyle" VALUE="0">
										<PARAM NAME="AllowRowSizing" VALUE="0">
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
										<PARAM NAME="CellNavigation" VALUE="0">
										<PARAM NAME="MaxSelectedRows" VALUE="0">
										<PARAM NAME="HeadStyleSet" VALUE="">
										<PARAM NAME="StyleSet" VALUE="">
										<PARAM NAME="ForeColorEven" VALUE="0">
										<PARAM NAME="ForeColorOdd" VALUE="0">
										<PARAM NAME="BackColorEven" VALUE="0">
										<PARAM NAME="BackColorOdd" VALUE="0">
										<PARAM NAME="Levels" VALUE="1">
										<PARAM NAME="RowHeight" VALUE="503">
										<PARAM NAME="ExtraHeight" VALUE="0">
										<PARAM NAME="ActiveRowStyleSet" VALUE="">
										<PARAM NAME="CaptionAlignment" VALUE="2">
										<PARAM NAME="SplitterPos" VALUE="0">
										<PARAM NAME="SplitterVisible" VALUE="0">
										<PARAM NAME="Columns.Count" VALUE="6">
										<!--TO-->        
										<PARAM NAME="Columns(0).Width" VALUE="850">
										<PARAM NAME="Columns(0).Visible" VALUE="-1">
										<PARAM NAME="Columns(0).Columns.Count" VALUE="1">
										<PARAM NAME="Columns(0).Caption" VALUE="To">
										<PARAM NAME="Columns(0).Name" VALUE="TO">
										<PARAM NAME="Columns(0).Alignment" VALUE="0">
										<PARAM NAME="Columns(0).CaptionAlignment" VALUE="2">
										<PARAM NAME="Columns(0).Bound" VALUE="0">
										<PARAM NAME="Columns(0).AllowSizing" VALUE="0">
										<PARAM NAME="Columns(0).DataField" VALUE="Column 0">
										<PARAM NAME="Columns(0).DataType" VALUE="8">
										<PARAM NAME="Columns(0).Level" VALUE="0">
										<PARAM NAME="Columns(0).NumberFormat" VALUE="">
										<PARAM NAME="Columns(0).Case" VALUE="0">
										<PARAM NAME="Columns(0).FieldLen" VALUE="256">
										<PARAM NAME="Columns(0).VertScrollBar" VALUE="0">
										<PARAM NAME="Columns(0).Locked" VALUE="0">
										<PARAM NAME="Columns(0).Style" VALUE="2">
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
										<!--CC-->         
										<PARAM NAME="Columns(1).Width" VALUE="850">
										<PARAM NAME="Columns(1).Visible" VALUE="-1">
										<PARAM NAME="Columns(1).Columns.Count" VALUE="1">
										<PARAM NAME="Columns(1).Caption" VALUE="Cc">
										<PARAM NAME="Columns(1).Name" VALUE="CC">
										<PARAM NAME="Columns(1).Alignment" VALUE="0">
										<PARAM NAME="Columns(1).CaptionAlignment" VALUE="3">
										<PARAM NAME="Columns(1).Bound" VALUE="0">
										<PARAM NAME="Columns(1).AllowSizing" VALUE="0">
										<PARAM NAME="Columns(1).DataField" VALUE="Column 1">
										<PARAM NAME="Columns(1).DataType" VALUE="8">
										<PARAM NAME="Columns(1).Level" VALUE="0">
										<PARAM NAME="Columns(1).NumberFormat" VALUE="">
										<PARAM NAME="Columns(1).Case" VALUE="0">
										<PARAM NAME="Columns(1).FieldLen" VALUE="256">
										<PARAM NAME="Columns(1).VertScrollBar" VALUE="0">
										<PARAM NAME="Columns(1).Locked" VALUE="0">
										<PARAM NAME="Columns(1).Style" VALUE="2">
										<PARAM NAME="Columns(1).ButtonsAlways" VALUE="0">
										<PARAM NAME="Columns(1).RowCount" VALUE="0">
										<PARAM NAME="Columns(1).ColCount" VALUE="1">
										<PARAM NAME="Columns(1).HasHeadForeColor" VALUE="0">
										<PARAM NAME="Columns(1).HasHeadBackColor" VALUE="0">
										<PARAM NAME="Columns(1).HasForeColor" VALUE="0">
										<PARAM NAME="Columns(1).HasBackColor" VALUE="0">
										<PARAM NAME="Columns(1).HeadForeColor" VALUE="0">
										<PARAM NAME="Columns(1).HeadBackColor" VALUE="0">
										<PARAM NAME="Columns(1).ForeColor" VALUE="0">
										<PARAM NAME="Columns(1).BackColor" VALUE="0">
										<PARAM NAME="Columns(1).HeadStyleSet" VALUE="">
										<PARAM NAME="Columns(1).StyleSet" VALUE="">
										<PARAM NAME="Columns(1).Nullable" VALUE="1">
										<PARAM NAME="Columns(1).Mask" VALUE="">
										<PARAM NAME="Columns(1).PromptInclude" VALUE="0">
										<PARAM NAME="Columns(1).ClipMode" VALUE="0">
										<PARAM NAME="Columns(1).PromptChar" VALUE="95">
										<!--BCC-->         
										<PARAM NAME="Columns(2).Width" VALUE="850">
										<PARAM NAME="Columns(2).Visible" VALUE="-1">
										<PARAM NAME="Columns(2).Columns.Count" VALUE="1">
										<PARAM NAME="Columns(2).Caption" VALUE="Bcc">
										<PARAM NAME="Columns(2).Name" VALUE="BCC">
										<PARAM NAME="Columns(2).Alignment" VALUE="0">
										<PARAM NAME="Columns(2).CaptionAlignment" VALUE="3">
										<PARAM NAME="Columns(2).Bound" VALUE="0">
										<PARAM NAME="Columns(2).AllowSizing" VALUE="0">
										<PARAM NAME="Columns(2).DataField" VALUE="Column 2">
										<PARAM NAME="Columns(2).DataType" VALUE="8">
										<PARAM NAME="Columns(2).Level" VALUE="0">
										<PARAM NAME="Columns(2).NumberFormat" VALUE="">
										<PARAM NAME="Columns(2).Case" VALUE="0">
										<PARAM NAME="Columns(2).FieldLen" VALUE="256">
										<PARAM NAME="Columns(2).VertScrollBar" VALUE="0">
										<PARAM NAME="Columns(2).Locked" VALUE="0">
										<PARAM NAME="Columns(2).Style" VALUE="2">
										<PARAM NAME="Columns(2).ButtonsAlways" VALUE="0">
										<PARAM NAME="Columns(2).Row.Count" VALUE="0">
										<PARAM NAME="Columns(2).Col.Count" VALUE="1">
										<PARAM NAME="Columns(2).HasHeadForeColor" VALUE="0">
										<PARAM NAME="Columns(2).HasHeadBackColor" VALUE="0">
										<PARAM NAME="Columns(2).HasForeColor" VALUE="0">
										<PARAM NAME="Columns(2).HasBackColor" VALUE="0">
										<PARAM NAME="Columns(2).HeadForeColor" VALUE="0">
										<PARAM NAME="Columns(2).HeadBackColor" VALUE="0">
										<PARAM NAME="Columns(2).ForeColor" VALUE="0">
										<PARAM NAME="Columns(2).BackColor" VALUE="0">
										<PARAM NAME="Columns(2).HeadStyleSet" VALUE="">
										<PARAM NAME="Columns(2).StyleSet" VALUE="">
										<PARAM NAME="Columns(2).Nullable" VALUE="1">
										<PARAM NAME="Columns(2).Mask" VALUE="">
										<PARAM NAME="Columns(2).PromptInclude" VALUE="0">
										<PARAM NAME="Columns(2).ClipMode" VALUE="0">
										<PARAM NAME="Columns(2).PromptChar" VALUE="95">
										<!--Recipient-->         
										<PARAM NAME="Columns(3).Width" VALUE="4875">
										<PARAM NAME="Columns(3).Visible" VALUE="-1">
										<PARAM NAME="Columns(3).Columns.Count" VALUE="1">
										<PARAM NAME="Columns(3).Caption" VALUE="Recipient">
										<PARAM NAME="Columns(3).Name" VALUE="Recipient">
										<PARAM NAME="Columns(3).Alignment" VALUE="0">
										<PARAM NAME="Columns(3).CaptionAlignment" VALUE="3">
										<PARAM NAME="Columns(3).Bound" VALUE="0">
										<PARAM NAME="Columns(3).AllowSizing" VALUE="1">
										<PARAM NAME="Columns(3).DataField" VALUE="Column 3">
										<PARAM NAME="Columns(3).DataType" VALUE="8">
										<PARAM NAME="Columns(3).Level" VALUE="0">
										<PARAM NAME="Columns(3).NumberFormat" VALUE="">
										<PARAM NAME="Columns(3).Case" VALUE="0">
										<PARAM NAME="Columns(3).FieldLen" VALUE="256">
										<PARAM NAME="Columns(3).VertScrollBar" VALUE="0">
										<PARAM NAME="Columns(3).Locked" VALUE="1">
										<PARAM NAME="Columns(3).Style" VALUE="0">
										<PARAM NAME="Columns(3).ButtonsAlways" VALUE="0">
										<PARAM NAME="Columns(3).RowCount" VALUE="0">
										<PARAM NAME="Columns(3).ColCount" VALUE="1">
										<PARAM NAME="Columns(3).HasHeadForeColor" VALUE="0">
										<PARAM NAME="Columns(3).HasHeadBackColor" VALUE="0">
										<PARAM NAME="Columns(3).HasForeColor" VALUE="0">
										<PARAM NAME="Columns(3).HasBackColor" VALUE="0">
										<PARAM NAME="Columns(3).HeadForeColor" VALUE="0">
										<PARAM NAME="Columns(3).HeadBackColor" VALUE="0">
										<PARAM NAME="Columns(3).ForeColor" VALUE="0">
										<PARAM NAME="Columns(3).BackColor" VALUE="0">
										<PARAM NAME="Columns(3).HeadStyleSet" VALUE="">
										<PARAM NAME="Columns(3).StyleSet" VALUE="">
										<PARAM NAME="Columns(3).Nullable" VALUE="1">
										<PARAM NAME="Columns(3).Mask" VALUE="">
										<PARAM NAME="Columns(3).PromptInclude" VALUE="0">
										<PARAM NAME="Columns(3).ClipMode" VALUE="0">
										<PARAM NAME="Columns(3).PromptChar" VALUE="95">
										<!--EmailID-->         
										<PARAM NAME="Columns(4).Width" VALUE="1000">
										<PARAM NAME="Columns(4).Visible" VALUE="0">
										<PARAM NAME="Columns(4).Columns.Count" VALUE="1">
										<PARAM NAME="Columns(4).Caption" VALUE="EmailID">
										<PARAM NAME="Columns(4).Name" VALUE="EmailID">
										<PARAM NAME="Columns(4).Alignment" VALUE="0">
										<PARAM NAME="Columns(4).CaptionAlignment" VALUE="3">
										<PARAM NAME="Columns(4).Bound" VALUE="0">
										<PARAM NAME="Columns(4).AllowSizing" VALUE="1">
										<PARAM NAME="Columns(4).DataField" VALUE="Column 4">
										<PARAM NAME="Columns(4).DataType" VALUE="8">
										<PARAM NAME="Columns(4).Level" VALUE="0">
										<PARAM NAME="Columns(4).NumberFormat" VALUE="">
										<PARAM NAME="Columns(4).Case" VALUE="0">
										<PARAM NAME="Columns(4).FieldLen" VALUE="256">
										<PARAM NAME="Columns(4).VertScrollBar" VALUE="0">
										<PARAM NAME="Columns(4).Locked" VALUE="0">
										<PARAM NAME="Columns(4).Style" VALUE="0">
										<PARAM NAME="Columns(4).ButtonsAlways" VALUE="0">
										<PARAM NAME="Columns(4).RowCount" VALUE="0">
										<PARAM NAME="Columns(4).ColCount" VALUE="1">
										<PARAM NAME="Columns(4).HasHeadForeColor" VALUE="0">
										<PARAM NAME="Columns(4).HasHeadBackColor" VALUE="0">
										<PARAM NAME="Columns(4).HasForeColor" VALUE="0">
										<PARAM NAME="Columns(4).HasBackColor" VALUE="0">
										<PARAM NAME="Columns(4).HeadForeColor" VALUE="0">
										<PARAM NAME="Columns(4).HeadBackColor" VALUE="0">
										<PARAM NAME="Columns(4).ForeColor" VALUE="0">
										<PARAM NAME="Columns(4).BackColor" VALUE="0">
										<PARAM NAME="Columns(4).HeadStyleSet" VALUE="">
										<PARAM NAME="Columns(4).StyleSet" VALUE="">
										<PARAM NAME="Columns(4).Nullable" VALUE="1">
										<PARAM NAME="Columns(4).Mask" VALUE="">
										<PARAM NAME="Columns(4).PromptInclude" VALUE="0">
										<PARAM NAME="Columns(4).ClipMode" VALUE="0">
										<PARAM NAME="Columns(4).PromptChar" VALUE="95">
										
										<!--EmailAddresses-->         
										<PARAM NAME="Columns(5).Width" VALUE="1000">
										<PARAM NAME="Columns(5).Visible" VALUE="0">
										<PARAM NAME="Columns(5).Columns.Count" VALUE="1">
										<PARAM NAME="Columns(5).Caption" VALUE="EmailAddresses">
										<PARAM NAME="Columns(5).Name" VALUE="EmailAddresses">
										<PARAM NAME="Columns(5).Alignment" VALUE="0">
										<PARAM NAME="Columns(5).CaptionAlignment" VALUE="3">
										<PARAM NAME="Columns(5).Bound" VALUE="0">
										<PARAM NAME="Columns(5).AllowSizing" VALUE="1">
										<PARAM NAME="Columns(5).DataField" VALUE="Column 5">
										<PARAM NAME="Columns(5).DataType" VALUE="8">
										<PARAM NAME="Columns(5).Level" VALUE="0">
										<PARAM NAME="Columns(5).NumberFormat" VALUE="">
										<PARAM NAME="Columns(5).Case" VALUE="0">
										<PARAM NAME="Columns(5).FieldLen" VALUE="256">
										<PARAM NAME="Columns(5).VertScrollBar" VALUE="0">
										<PARAM NAME="Columns(5).Locked" VALUE="0">
										<PARAM NAME="Columns(5).Style" VALUE="0">
										<PARAM NAME="Columns(5).ButtonsAlways" VALUE="0">
										<PARAM NAME="Columns(5).RowCount" VALUE="0">
										<PARAM NAME="Columns(5).ColCount" VALUE="1">
										<PARAM NAME="Columns(5).HasHeadForeColor" VALUE="0">
										<PARAM NAME="Columns(5).HasHeadBackColor" VALUE="0">
										<PARAM NAME="Columns(5).HasForeColor" VALUE="0">
										<PARAM NAME="Columns(5).HasBackColor" VALUE="0">
										<PARAM NAME="Columns(5).HeadForeColor" VALUE="0">
										<PARAM NAME="Columns(5).HeadBackColor" VALUE="0">
										<PARAM NAME="Columns(5).ForeColor" VALUE="0">
										<PARAM NAME="Columns(5).BackColor" VALUE="0">
										<PARAM NAME="Columns(5).HeadStyleSet" VALUE="">
										<PARAM NAME="Columns(5).StyleSet" VALUE="">
										<PARAM NAME="Columns(5).Nullable" VALUE="1">
										<PARAM NAME="Columns(5).Mask" VALUE="">
										<PARAM NAME="Columns(5).PromptInclude" VALUE="0">
										<PARAM NAME="Columns(5).ClipMode" VALUE="0">
										<PARAM NAME="Columns(5).PromptChar" VALUE="95">
													
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
								</td>
								<td width=80 valign=bottom>
									<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD width=10>
												<INPUT id=cmdOK type=button value=OK name=cmdOK class="btn" style="WIDTH: 80px" width="80" 
												    onclick="emailEvent();"
					                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                                        onfocus="try{button_onFocus(this);}catch(e){}"
			                                        onblur="try{button_onBlur(this);}catch(e){}" />
											</TD>
										</TR>
										<TR height=10>
											<TD>
											</TD>
										</TR>
										<TR>
											<TD width=10>
												<INPUT id=cmdCancel type=button class="btn" value="Cancel" name=cmdCancel style="WIDTH: 80px" width="80"
												    onclick="cancelClick();"
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
	</tr> 
</TABLE>
</form>

<FORM action="default_Submit" method=post id="frmGoto" name=frmGoto style="visibility:hidden;display:none">
    <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</FORM>

<FORM id=frmUseful name=frmUseful style="visibility:hidden;display:none">
	<INPUT type="hidden" id=txtUserName name=txtUserName value="<%=session("username")%>">
<%
    Dim cmdDefinition
    Dim prmModuleKey
    Dim prmParameterValue
    Dim prmParameterKey
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

<FORM name=frmList id=frmList style="visibility:hidden;display:none">

<%
	'Get the required Email information
	dim cmdEmail
	dim rsEmail
	dim i 
	dim sAddline
	dim sEmailAddresses
	dim cmdEmailAddr
	dim rstEmailAddr
    Dim iLoop
    Dim prmEmailGroupID
	
	i = 0
	iLoop = 0 
	sAddline = vbnullString 
	sEmailAddresses = vbNullString 
	
    cmdEmail = CreateObject("ADODB.Command")
	cmdEmail.CommandText = "spASRIntGetEventLogEmails"
	cmdEmail.CommandType = 4 'Stored Procedure
    cmdEmail.ActiveConnection = Session("databaseConnection")
	
    Err.Clear()
    rsEmail = cmdEmail.Execute
	
	if not (rsEmail.bof and rsEmail.eof) then
		
		do until rsEmail.eof
			i = i + 1
			sEmailAddresses = vbNullString 
			sAddline = vbNullString 
			sAddline = "0" & vbTab & "0" & vbTab & "0" & vbTab
            sAddline = sAddline & rsEmail.Fields("Name").Value & vbTab
            sAddline = sAddline & rsEmail.Fields("EmailGroupID").Value & vbTab
			
            If rsEmail.Fields("EmailGroupID").Value < 1 Then
                sAddline = sAddline & rsEmail.Fields("Name").value
            Else

                cmdEmailAddr = CreateObject("ADODB.Command")
                cmdEmailAddr.CommandText = "spASRIntGetEmailGroupAddresses"
                cmdEmailAddr.CommandType = 4 ' Stored procedure
                cmdEmailAddr.ActiveConnection = Session("databaseConnection")

                prmEmailGroupID = cmdEmailAddr.CreateParameter("EmailGroupID", 3, 1) ' 3=integer, 1=input
                cmdEmailAddr.Parameters.Append(prmEmailGroupID)
                prmEmailGroupID.value = CleanNumeric(rsEmail.Fields("EmailGroupID").Value)

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
				
                sAddline = sAddline & sEmailAddresses
            End If
			
            Response.Write("<INPUT type=hidden name=txtEmailGroup_" & i & " id=txtEmailGroup_" & i & " value=""" & sAddline & """>" & vbCrLf)
			
			rsEmail.movenext
		loop
	end if
	
    rsEmail = Nothing
    cmdEmail = Nothing
%>

</FORM>

<FORM name=frmEmailDetails id=frmEmailDetails style="visibility:hidden;display:none">

<%
	'Get the required Email information
	dim cmdEmailDetails
	dim rsEmailDetails
	dim sEmailInfo
	dim iLastEventID
	dim iDetailCount
	
    Dim objUtilities
    Dim prmSelectedIDs
    Dim prmSubject
    Dim prmEmailOrderColumn
    Dim prmEmailOrderOrder
		
    objUtilities = Session("UtilitiesObject")
		
    cmdEmailDetails = CreateObject("ADODB.Command")
	cmdEmailDetails.CommandText = "spASRIntGetEventLogEmailInfo"
	cmdEmailDetails.CommandType = 4 'Stored Procedure
    cmdEmailDetails.ActiveConnection = Session("databaseConnection")
	
    prmSelectedIDs = cmdEmailDetails.CreateParameter("selectedids", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdEmailDetails.Parameters.Append(prmSelectedIDs)
    prmSelectedIDs.value = Request("txtSelectedEventIDs")

    prmSubject = cmdEmailDetails.CreateParameter("subject", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
    cmdEmailDetails.Parameters.Append(prmSubject)
	
    prmEmailOrderColumn = cmdEmailDetails.CreateParameter("emailOrderColumn", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdEmailDetails.Parameters.Append(prmEmailOrderColumn)
    prmEmailOrderColumn.value = CStr(Request("txtEmailOrderColumn"))
		
    prmEmailOrderOrder = cmdEmailDetails.CreateParameter("emailOrderOrder", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdEmailDetails.Parameters.Append(prmEmailOrderOrder)
    prmEmailOrderOrder.value = CStr(Request("txtEmailOrderOrder"))
	
	
    Err.Clear()
    rsEmailDetails = cmdEmailDetails.Execute
	
	sEmailInfo = vbNullString 
	iDetailCount = 0 
	iLastEventID = -1
	
	dim EventCounter
	EventCounter = 0 
	
    If (Err.Number <> 0) Then
        sErrorDescription = "Error getting the event log records." & vbCrLf & formatError(Err.Description)
    End If

	if len(sErrorDescription) = 0 then
		if not (rsEmailDetails.bof and rsEmailDetails.eof) then
			do until rsEmailDetails.eof
			
                If iLastEventID <> rsEmailDetails.Fields("ID").value Then
					
                    EventCounter = EventCounter + 1
                    Response.Write(CStr(EventCounter))

                    sEmailInfo = sEmailInfo & StrDup(Len(rsEmailDetails.Fields("Name").Value) + 30, "-") & vbCrLf
                    sEmailInfo = sEmailInfo & "Event Name : " & rsEmailDetails.Fields("Name").Value & vbCrLf
                    sEmailInfo = sEmailInfo & StrDup(Len(rsEmailDetails.Fields("Name").Value) + 30, "-") & vbCrLf
					
                    sEmailInfo = sEmailInfo & "Mode :		" & rsEmailDetails.Fields("Mode").Value & vbCrLf & vbCrLf
					
                    sEmailInfo = sEmailInfo & "Start Time :	" & ConvertSqlDateToLocale(rsEmailDetails.Fields("DateTime").Value) & " " & ConvertSqlDateToTime(rsEmailDetails.Fields("DateTime").Value) & vbCrLf
                    If IsDBNull(rsEmailDetails.Fields("EndTime").Value) Then
                        sEmailInfo = sEmailInfo & "End Time :	" & vbCrLf
                    Else
                        sEmailInfo = sEmailInfo & "End Time :	" & ConvertSqlDateToLocale(rsEmailDetails.Fields("DateTime").Value) & " " & ConvertSqlDateToTime(rsEmailDetails.Fields("EndTime").Value) & vbCrLf
                    End If
                    sEmailInfo = sEmailInfo & "Duration :	" & objUtilities.FormatEventDuration(CLng(rsEmailDetails.Fields("Duration").Value)) & vbCrLf
					
                    sEmailInfo = sEmailInfo & "Type :		" & rsEmailDetails.Fields("Type").Value & vbCrLf
                    sEmailInfo = sEmailInfo & "Status :		" & rsEmailDetails.Fields("Status").Value & vbCrLf
                    sEmailInfo = sEmailInfo & "User name :	" & rsEmailDetails.Fields("Username").Value & vbCrLf & vbCrLf
					
                    If Request("txtFromMain") = 0 Then
                        If Request("txtBatchy") Then
                            sEmailInfo = sEmailInfo & Request("txtBatchInfo") & vbCrLf
                        End If
                    Else
                        If (Not IsDBNull(rsEmailDetails.Fields("BatchName"))) And (Len(rsEmailDetails.Fields("BatchName").Value) > 0) Then
                            sEmailInfo = sEmailInfo & "Batch Job Name	: " & rsEmailDetails.Fields("BatchName").Value & vbCrLf & vbCrLf
                        End If
                    End If
										
                    sEmailInfo = sEmailInfo & "Records Successful :	" & rsEmailDetails.Fields("SuccessCount").Value & vbCrLf
                    sEmailInfo = sEmailInfo & "Records Failed :		" & rsEmailDetails.Fields("FailCount").Value & vbCrLf & vbCrLf
					
                    sEmailInfo = sEmailInfo & "Details : " & vbCrLf & vbCrLf
					
                    iLastEventID = rsEmailDetails.Fields("ID").Value
                    iDetailCount = 0
                End If
				
				iDetailCount = iDetailCount + 1
				
                If rsEmailDetails.Fields("count").Value > 0 Then
                    If (Not IsDBNull(rsEmailDetails.Fields("Notes"))) And (Len(rsEmailDetails.Fields("Notes").Value) > 0) Then
                        sEmailInfo = sEmailInfo & "*** Log Entry " & CStr(iDetailCount) & " of " & CStr(rsEmailDetails.Fields("count").Value) & " ***" & vbCrLf
                        sEmailInfo = sEmailInfo & rsEmailDetails.Fields("Notes").Value
                    End If
                Else
                    sEmailInfo = sEmailInfo & "There are no details for this event log entry" & vbCrLf
                End If
				
				sEmailInfo = sEmailInfo	& vbCrLf & vbCrLf & vbCrLf
				
				rsEmailDetails.Movenext
			loop
		
            Response.Write("<INPUT type=hidden name=txtEventDeleted id=txtEventDeleted value=0>" & vbCrLf)
			
        Else
            Response.Write("<INPUT type=hidden name=txtEventDeleted id=txtEventDeleted value=1>" & vbCrLf)
		end if
	end if
	
	rsEmailDetails.close
    rsEmailDetails = Nothing

    Response.Write("<INPUT type=hidden name=txtBody id=txtBody value=""" & Replace(sEmailInfo, """", "&quot;") & """>" & vbCrLf)
    Response.Write("<INPUT type=hidden name=txtSubject id=txtSubject value=""" & Replace(cmdEmailDetails.Parameters("subject").Value, """", "&quot;") & """>" & vbCrLf)
	
    cmdEmailDetails = Nothing
    objUtilities = Nothing
	
%>

</FORM>
    
    
    
<script type="text/javascript">
    function emailSelection_window_onload() {


        if (frmEmailDetails.txtEventDeleted.value == 1) {
            OpenHR.messageBox("This record no longer exists in the event log.", 48, "Event Log");
            
            try {               
                window.dialogArguments.parent.frames("workframe").refreshGrid();
            } catch(e) {
            }

            self.close();
        } else {
            setGridFont(frmEmail.ssOleDBGridEmail);

            populateEmailList();
        }
    }
</script>

<script type="text/javascript" id=scptGeneralFunctions>
<!--

    function populateEmailList()
    {
        var sAddLine = ''
	
        frmEmail.ssOleDBGridEmail.focus();
        frmEmail.ssOleDBGridEmail.Redraw = false;
	
        for (var i=0; i<frmList.elements.length; i++)
        {
            sAddLine = frmList.elements[i].value;
            frmEmail.ssOleDBGridEmail.AddItem(sAddLine);
        }
		
        frmEmail.ssOleDBGridEmail.Redraw = true;
    }
	
    function emailEvent()
    {
        var bOK = false;
        var sTo = getEmailList(0);
        var sCC = getEmailList(1);
        var sBCC = getEmailList(2);
        var sSubject = getSubject();
        var sBody = getBody();

        bOK = OpenHR.sendMail(sTo,sSubject,sBody,sCC,sBCC);
        
      //  OpenHR.SendMail()
        

	
        self.close();
        return bOK;
    }
	
    function getEmailList(iSendType)
    {
        var sEmailList = '';
	
        frmEmail.ssOleDBGridEmail.Redraw = false;
        frmEmail.ssOleDBGridEmail.MoveFirst();
        for (var i=0; i < frmEmail.ssOleDBGridEmail.Rows; i++)
        {
            if (frmEmail.ssOleDBGridEmail.Columns(iSendType).value == -1)
            {
                if (sEmailList.length > 0)
                {
                    sEmailList = sEmailList + '; ';
                }
                sEmailList = sEmailList + frmEmail.ssOleDBGridEmail.Columns("EmailAddresses").Text;
            }
            frmEmail.ssOleDBGridEmail.MoveNext();
        }
        frmEmail.ssOleDBGridEmail.Redraw = true;
		
        return (sEmailList);
    }

    function getSubject()
    {
        return frmEmailDetails.txtSubject.value;
    }

    function getBody()
    {
        return frmEmailDetails.txtBody.value;
    }

    function cancelClick()
    {
        self.close();
    }
	
    -->
</script>
    
<script type="text/javascript">
    emailSelection_window_onload();
</script>


</body>
</html>
