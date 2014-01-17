<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<!DOCTYPE html>

<html>
<head runat="server">
	<title>Event Log Selection - OpenHR Intranet</title>
	<script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/OpenHR_ActiveX")%>" type="text/javascript"></script>
	
	<%--Here's the stylesheets for the font-icons displayed on the dashboard for wireframe and tile layouts--%>
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/fonts/SSI80v194934/style.css")%>" rel="stylesheet" />

	<%--Base stylesheets--%>
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />

	<%--stylesheet for slide-out dmi menu--%>
	<link href="<%: Url.LatestContent("~/Content/contextmenustyle.css")%>" rel="stylesheet" type="text/css" />

	<%--ThemeRoller stylesheet--%>
	<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />

	<%--jQuery Grid Stylesheet--%>
	<link href="<%: Url.LatestContent("~/Content/ui.jqgrid.css")%>" rel="stylesheet" type="text/css" />


	

</head>
<body>

	<object
		classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
		id="Microsoft_Licensed_Class_Manager_1_0"
		viewastext>
		<param name="LPKPath" value="lpks/main.lpk">
	</object>

	<form id="frmEmail">
		<table >
			<tr>
				<td>
					<table style="width: 100%; height: 100%" class="invisible">
						<tr style="height: 5px">
							<td colspan="3"></td>
						</tr>
						<tr>
							
							<td width="5"></td>
							<td>
								<table class="invisible"
									style="border-spacing: 4px; padding: 10px; width: 100%; height: 100%; float: left">
									<tr>
										<td width="100%">
											<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
												codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
												id="ssOleDBGridEmail"
												name="ssOleDBGridEmail"
												style="HEIGHT: 250px; VISIBILITY: visible; WIDTH: 99%">
												<param name="ScrollBars" value="0">
												<param name="_Version" value="196617">
												<param name="DataMode" value="2">
												<param name="Cols" value="0">
												<param name="Rows" value="0">
												<param name="BorderStyle" value="1">
												<param name="RecordSelectors" value="0">
												<param name="GroupHeaders" value="-1">
												<param name="ColumnHeaders" value="-1">
												<param name="GroupHeadLines" value="1">
												<param name="HeadLines" value="2">
												<param name="FieldDelimiter" value="(None)">
												<param name="FieldSeparator" value="(Tab)">
												<param name="Row.Count" value="0">
												<param name="Col.Count" value="6">
												<param name="stylesets.count" value="0">
												<param name="TagVariant" value="EMPTY">
												<param name="UseGroups" value="0">
												<param name="HeadFont3D" value="0">
												<param name="Font3D" value="0">
												<param name="DividerType" value="3">
												<param name="DividerStyle" value="1">
												<param name="DefColWidth" value="0">
												<param name="BeveColorScheme" value="2">
												<param name="BevelColorFrame" value="0">
												<param name="BevelColorHighlight" value="0">
												<param name="BevelColorShadow" value="0">
												<param name="BevelColorFace" value="0">
												<param name="CheckBox3D" value="0">
												<param name="AllowAddNew" value="0">
												<param name="AllowDelete" value="0">
												<param name="AllowUpdate" value="-1">
												<param name="MultiLine" value="0">
												<param name="ActiveCellStyleSet" value="">
												<param name="RowSelectionStyle" value="0">
												<param name="AllowRowSizing" value="0">
												<param name="AllowGroupSizing" value="0">
												<param name="AllowColumnSizing" value="-1">
												<param name="AllowGroupMoving" value="0">
												<param name="AllowColumnMoving" value="0">
												<param name="AllowGroupSwapping" value="0">
												<param name="AllowColumnSwapping" value="0">
												<param name="AllowGroupShrinking" value="0">
												<param name="AllowColumnShrinking" value="0">
												<param name="AllowDragDrop" value="0">
												<param name="UseExactRowCount" value="-1">
												<param name="SelectTypeCol" value="0">
												<param name="SelectTypeRow" value="1">
												<param name="SelectByCell" value="-1">
												<param name="BalloonHelp" value="0">
												<param name="RowNavigation" value="1">
												<param name="CellNavigation" value="0">
												<param name="MaxSelectedRows" value="0">
												<param name="HeadStyleSet" value="">
												<param name="StyleSet" value="">
												<param name="ForeColorEven" value="0">
												<param name="ForeColorOdd" value="0">
												<param name="BackColorEven" value="0">
												<param name="BackColorOdd" value="0">
												<param name="Levels" value="1">
												<param name="RowHeight" value="609">
												<param name="ExtraHeight" value="0">
												<param name="ActiveRowStyleSet" value="">
												<param name="CaptionAlignment" value="2">
												<param name="SplitterPos" value="0">
												<param name="SplitterVisible" value="0">
												<param name="Columns.Count" value="6">
												<!--TO-->
												<param name="Columns(0).Width" value="850">
												<param name="Columns(0).Visible" value="-1">
												<param name="Columns(0).Columns.Count" value="1">
												<param name="Columns(0).Caption" value="To">
												<param name="Columns(0).Name" value="TO">
												<param name="Columns(0).Alignment" value="0">
												<param name="Columns(0).CaptionAlignment" value="2">
												<param name="Columns(0).Bound" value="0">
												<param name="Columns(0).AllowSizing" value="0">
												<param name="Columns(0).DataField" value="Column 0">
												<param name="Columns(0).DataType" value="8">
												<param name="Columns(0).Level" value="0">
												<param name="Columns(0).NumberFormat" value="">
												<param name="Columns(0).Case" value="0">
												<param name="Columns(0).FieldLen" value="256">
												<param name="Columns(0).VertScrollBar" value="0">
												<param name="Columns(0).Locked" value="0">
												<param name="Columns(0).Style" value="2">
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
												<!--CC-->
												<param name="Columns(1).Width" value="850">
												<param name="Columns(1).Visible" value="-1">
												<param name="Columns(1).Columns.Count" value="1">
												<param name="Columns(1).Caption" value="Cc">
												<param name="Columns(1).Name" value="CC">
												<param name="Columns(1).Alignment" value="0">
												<param name="Columns(1).CaptionAlignment" value="3">
												<param name="Columns(1).Bound" value="0">
												<param name="Columns(1).AllowSizing" value="0">
												<param name="Columns(1).DataField" value="Column 1">
												<param name="Columns(1).DataType" value="8">
												<param name="Columns(1).Level" value="0">
												<param name="Columns(1).NumberFormat" value="">
												<param name="Columns(1).Case" value="0">
												<param name="Columns(1).FieldLen" value="256">
												<param name="Columns(1).VertScrollBar" value="0">
												<param name="Columns(1).Locked" value="0">
												<param name="Columns(1).Style" value="2">
												<param name="Columns(1).ButtonsAlways" value="0">
												<param name="Columns(1).RowCount" value="0">
												<param name="Columns(1).ColCount" value="1">
												<param name="Columns(1).HasHeadForeColor" value="0">
												<param name="Columns(1).HasHeadBackColor" value="0">
												<param name="Columns(1).HasForeColor" value="0">
												<param name="Columns(1).HasBackColor" value="0">
												<param name="Columns(1).HeadForeColor" value="0">
												<param name="Columns(1).HeadBackColor" value="0">
												<param name="Columns(1).ForeColor" value="0">
												<param name="Columns(1).BackColor" value="0">
												<param name="Columns(1).HeadStyleSet" value="">
												<param name="Columns(1).StyleSet" value="">
												<param name="Columns(1).Nullable" value="1">
												<param name="Columns(1).Mask" value="">
												<param name="Columns(1).PromptInclude" value="0">
												<param name="Columns(1).ClipMode" value="0">
												<param name="Columns(1).PromptChar" value="95">
												<!--BCC-->
												<param name="Columns(2).Width" value="850">
												<param name="Columns(2).Visible" value="-1">
												<param name="Columns(2).Columns.Count" value="1">
												<param name="Columns(2).Caption" value="Bcc">
												<param name="Columns(2).Name" value="BCC">
												<param name="Columns(2).Alignment" value="0">
												<param name="Columns(2).CaptionAlignment" value="3">
												<param name="Columns(2).Bound" value="0">
												<param name="Columns(2).AllowSizing" value="0">
												<param name="Columns(2).DataField" value="Column 2">
												<param name="Columns(2).DataType" value="8">
												<param name="Columns(2).Level" value="0">
												<param name="Columns(2).NumberFormat" value="">
												<param name="Columns(2).Case" value="0">
												<param name="Columns(2).FieldLen" value="256">
												<param name="Columns(2).VertScrollBar" value="0">
												<param name="Columns(2).Locked" value="0">
												<param name="Columns(2).Style" value="2">
												<param name="Columns(2).ButtonsAlways" value="0">
												<param name="Columns(2).Row.Count" value="0">
												<param name="Columns(2).Col.Count" value="1">
												<param name="Columns(2).HasHeadForeColor" value="0">
												<param name="Columns(2).HasHeadBackColor" value="0">
												<param name="Columns(2).HasForeColor" value="0">
												<param name="Columns(2).HasBackColor" value="0">
												<param name="Columns(2).HeadForeColor" value="0">
												<param name="Columns(2).HeadBackColor" value="0">
												<param name="Columns(2).ForeColor" value="0">
												<param name="Columns(2).BackColor" value="0">
												<param name="Columns(2).HeadStyleSet" value="">
												<param name="Columns(2).StyleSet" value="">
												<param name="Columns(2).Nullable" value="1">
												<param name="Columns(2).Mask" value="">
												<param name="Columns(2).PromptInclude" value="0">
												<param name="Columns(2).ClipMode" value="0">
												<param name="Columns(2).PromptChar" value="95">
												<!--Recipient-->
												<param name="Columns(3).Width" value="12000">
												<param name="Columns(3).Visible" value="-1">
												<param name="Columns(3).Columns.Count" value="1">
												<param name="Columns(3).Caption" value="Recipient">
												<param name="Columns(3).Name" value="Recipient">
												<param name="Columns(3).Alignment" value="0">
												<param name="Columns(3).CaptionAlignment" value="3">
												<param name="Columns(3).Bound" value="0">
												<param name="Columns(3).AllowSizing" value="1">
												<param name="Columns(3).DataField" value="Column 3">
												<param name="Columns(3).DataType" value="8">
												<param name="Columns(3).Level" value="0">
												<param name="Columns(3).NumberFormat" value="">
												<param name="Columns(3).Case" value="0">
												<param name="Columns(3).FieldLen" value="256">
												<param name="Columns(3).VertScrollBar" value="0">
												<param name="Columns(3).Locked" value="1">
												<param name="Columns(3).Style" value="0">
												<param name="Columns(3).ButtonsAlways" value="0">
												<param name="Columns(3).RowCount" value="0">
												<param name="Columns(3).ColCount" value="1">
												<param name="Columns(3).HasHeadForeColor" value="0">
												<param name="Columns(3).HasHeadBackColor" value="0">
												<param name="Columns(3).HasForeColor" value="0">
												<param name="Columns(3).HasBackColor" value="0">
												<param name="Columns(3).HeadForeColor" value="0">
												<param name="Columns(3).HeadBackColor" value="0">
												<param name="Columns(3).ForeColor" value="0">
												<param name="Columns(3).BackColor" value="0">
												<param name="Columns(3).HeadStyleSet" value="">
												<param name="Columns(3).StyleSet" value="">
												<param name="Columns(3).Nullable" value="1">
												<param name="Columns(3).Mask" value="">
												<param name="Columns(3).PromptInclude" value="0">
												<param name="Columns(3).ClipMode" value="0">
												<param name="Columns(3).PromptChar" value="95">
												<!--EmailID-->
												<param name="Columns(4).Width" value="1000">
												<param name="Columns(4).Visible" value="0">
												<param name="Columns(4).Columns.Count" value="1">
												<param name="Columns(4).Caption" value="EmailID">
												<param name="Columns(4).Name" value="EmailID">
												<param name="Columns(4).Alignment" value="0">
												<param name="Columns(4).CaptionAlignment" value="3">
												<param name="Columns(4).Bound" value="0">
												<param name="Columns(4).AllowSizing" value="1">
												<param name="Columns(4).DataField" value="Column 4">
												<param name="Columns(4).DataType" value="8">
												<param name="Columns(4).Level" value="0">
												<param name="Columns(4).NumberFormat" value="">
												<param name="Columns(4).Case" value="0">
												<param name="Columns(4).FieldLen" value="256">
												<param name="Columns(4).VertScrollBar" value="0">
												<param name="Columns(4).Locked" value="0">
												<param name="Columns(4).Style" value="0">
												<param name="Columns(4).ButtonsAlways" value="0">
												<param name="Columns(4).RowCount" value="0">
												<param name="Columns(4).ColCount" value="1">
												<param name="Columns(4).HasHeadForeColor" value="0">
												<param name="Columns(4).HasHeadBackColor" value="0">
												<param name="Columns(4).HasForeColor" value="0">
												<param name="Columns(4).HasBackColor" value="0">
												<param name="Columns(4).HeadForeColor" value="0">
												<param name="Columns(4).HeadBackColor" value="0">
												<param name="Columns(4).ForeColor" value="0">
												<param name="Columns(4).BackColor" value="0">
												<param name="Columns(4).HeadStyleSet" value="">
												<param name="Columns(4).StyleSet" value="">
												<param name="Columns(4).Nullable" value="1">
												<param name="Columns(4).Mask" value="">
												<param name="Columns(4).PromptInclude" value="0">
												<param name="Columns(4).ClipMode" value="0">
												<param name="Columns(4).PromptChar" value="95">

												<!--EmailAddresses-->
												<param name="Columns(5).Width" value="1000">
												<param name="Columns(5).Visible" value="0">
												<param name="Columns(5).Columns.Count" value="1">
												<param name="Columns(5).Caption" value="EmailAddresses">
												<param name="Columns(5).Name" value="EmailAddresses">
												<param name="Columns(5).Alignment" value="0">
												<param name="Columns(5).CaptionAlignment" value="3">
												<param name="Columns(5).Bound" value="0">
												<param name="Columns(5).AllowSizing" value="1">
												<param name="Columns(5).DataField" value="Column 5">
												<param name="Columns(5).DataType" value="8">
												<param name="Columns(5).Level" value="0">
												<param name="Columns(5).NumberFormat" value="">
												<param name="Columns(5).Case" value="0">
												<param name="Columns(5).FieldLen" value="256">
												<param name="Columns(5).VertScrollBar" value="0">
												<param name="Columns(5).Locked" value="0">
												<param name="Columns(5).Style" value="0">
												<param name="Columns(5).ButtonsAlways" value="0">
												<param name="Columns(5).RowCount" value="0">
												<param name="Columns(5).ColCount" value="1">
												<param name="Columns(5).HasHeadForeColor" value="0">
												<param name="Columns(5).HasHeadBackColor" value="0">
												<param name="Columns(5).HasForeColor" value="0">
												<param name="Columns(5).HasBackColor" value="0">
												<param name="Columns(5).HeadForeColor" value="0">
												<param name="Columns(5).HeadBackColor" value="0">
												<param name="Columns(5).ForeColor" value="0">
												<param name="Columns(5).BackColor" value="0">
												<param name="Columns(5).HeadStyleSet" value="">
												<param name="Columns(5).StyleSet" value="">
												<param name="Columns(5).Nullable" value="1">
												<param name="Columns(5).Mask" value="">
												<param name="Columns(5).PromptInclude" value="0">
												<param name="Columns(5).ClipMode" value="0">
												<param name="Columns(5).PromptChar" value="95">

												<param name="UseDefaults" value="-1">
												<param name="TabNavigation" value="1">
												<param name="BatchUpdate" value="0">
												<param name="_ExtentX" value="28019">
												<param name="_ExtentY" value="4974">
												<param name="_StockProps" value="79">
												<param name="Caption" value="">
												<param name="ForeColor" value="0">
												<param name="BackColor" value="0">
												<param name="Enabled" value="-1">
												<param name="DataMember" value="">
											</object>
										</td>
									</tr>
									<tr>
									<td>
											<%--<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>--%>
											<table class="invisible">
												<tr>
													<td>
														<input id="cmdOK" type="button" value="OK" name="cmdOK" style="WIDTH: 80px" class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br"
															onclick="emailEvent();"
															onmouseover="try{button_onMouseOver(this);}catch(e){}"
															onmouseout="try{button_onMouseOut(this);}catch(e){}"
															onfocus="try{button_onFocus(this);}catch(e){}"
															onblur="try{button_onBlur(this);}catch(e){}" />
													</td>

													<td>
														<input id="cmdCancel" type="button" value="Cancel" name="cmdCancel" style="WIDTH: 80px" class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br"
															onclick="cancelClick();"
															onmouseover="try{button_onMouseOver(this);}catch(e){}"
															onmouseout="try{button_onMouseOut(this);}catch(e){}"
															onfocus="try{button_onFocus(this);}catch(e){}"
															onblur="try{button_onBlur(this);}catch(e){}" />
													</td>
												</tr>
											</table>
										</td>
									</tr>
										
									</tr>
								</table>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</form>

	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
		<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
	</form>

	<form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
		<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
		<%

			Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)

			Dim sParameterValue As String = objDatabase.GetModuleParameter("MODULE_PERSONNEL", "Param_TablePersonnel")
			Response.Write("<input type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & sParameterValue & ">" & vbCrLf)
		
			Response.Write("<input type='hidden' id=txtErrorDescription name=txtErrorDescription value="""">" & vbCrLf)
			Response.Write("<input type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
			
			
		%>
	</form>

	<form name="frmList" id="frmList" style="visibility: hidden; display: none">

		<%

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
			
			'Get the required Email information
			Dim sErrorDescription As String
			Dim cmdEmail
			Dim rsEmail
			Dim i
			Dim sAddline
			Dim sEmailAddresses
			Dim iLoop
	
			i = 0
			iLoop = 0
			sAddline = vbNullString
			sEmailAddresses = vbNullString
	
			cmdEmail = CreateObject("ADODB.Command")
			cmdEmail.CommandText = "spASRIntGetEventLogEmails"
			cmdEmail.CommandType = 4 'Stored Procedure
			cmdEmail.ActiveConnection = Session("databaseConnection")
	
			Err.Clear()
			rsEmail = cmdEmail.Execute
	
			If Not (rsEmail.bof And rsEmail.eof) Then
		
				Do Until rsEmail.eof
					i = i + 1
					sEmailAddresses = vbNullString
					sAddline = vbNullString
					sAddline = "0" & vbTab & "0" & vbTab & "0" & vbTab
					sAddline = sAddline & rsEmail.Fields("Name").Value & vbTab
					sAddline = sAddline & rsEmail.Fields("EmailGroupID").Value & vbTab
			
					If rsEmail.Fields("EmailGroupID").Value < 1 Then
						sAddline = sAddline & rsEmail.Fields("Name").value
					Else

						Try
							Dim rstEmailAddr = objDataAccess.GetDataTable("spASRIntGetEmailGroupAddresses", CommandType.StoredProcedure _
										, New SqlParameter("EmailGroupID", SqlDbType.Int) With {.Value = rsEmail.Fields("EmailGroupID").Value})

							iLoop = 0
							If Not rstEmailAddr Is Nothing Then
								For Each objRow In rstEmailAddr.Rows
									
									If iLoop > 1 Then
										sEmailAddresses = sEmailAddresses & ";"
									End If
									
									sEmailAddresses = sEmailAddresses & objRow(0).ToString()
									iLoop += 1
								Next
							End If

						Catch ex As Exception
							sErrorDescription = "Error getting the email addresses for group." & vbCrLf & FormatError(ex.Message)
						End Try
									
						sAddline = sAddline & sEmailAddresses
					End If
			
					Response.Write("<INPUT type=hidden name=txtEmailGroup_" & i & " id=txtEmailGroup_" & i & " value=""" & sAddline & """>" & vbCrLf)
			
					rsEmail.movenext()
				Loop
			End If
	
			rsEmail = Nothing
			cmdEmail = Nothing
		%>
	</form>

	<form name="frmEmailDetails" id="frmEmailDetails" style="visibility: hidden; display: none; width:100%">

		<%
			'Get the required Email information
			Dim cmdEmailDetails
			Dim rsEmailDetails
			Dim sEmailInfo
			Dim iLastEventID
			Dim iDetailCount
	
			Dim objUtilities
			Dim prmSelectedIDs
			Dim prmSubject
			Dim prmEmailOrderColumn
			Dim prmEmailOrderOrder
		
			objUtilities = Session("UtilitiesObject")
		
			cmdEmailDetails = CreateObject("ADODB.Command")
			cmdEmailDetails.CommandText = "spASRIntGetEventLogEmailInfo"
			cmdEmailDetails.CommandType = 4	'Stored Procedure
			cmdEmailDetails.ActiveConnection = Session("databaseConnection")
	
			prmSelectedIDs = cmdEmailDetails.CreateParameter("selectedids", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdEmailDetails.Parameters.Append(prmSelectedIDs)
			prmSelectedIDs.value = Request("txtSelectedEventIDs")

			prmSubject = cmdEmailDetails.CreateParameter("subject", 200, 2, 8000)	' 200=varchar, 2=output, 8000=size
			cmdEmailDetails.Parameters.Append(prmSubject)
	
			prmEmailOrderColumn = cmdEmailDetails.CreateParameter("emailOrderColumn", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdEmailDetails.Parameters.Append(prmEmailOrderColumn)
			prmEmailOrderColumn.value = CStr(Request("txtEmailOrderColumn"))
		
			prmEmailOrderOrder = cmdEmailDetails.CreateParameter("emailOrderOrder", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdEmailDetails.Parameters.Append(prmEmailOrderOrder)
			prmEmailOrderOrder.value = CStr(Request("txtEmailOrderOrder"))
	
	
			Err.Clear()
			rsEmailDetails = cmdEmailDetails.Execute
	
			sEmailInfo = vbNullString
			iDetailCount = 0
			iLastEventID = -1
	
			Dim EventCounter
			EventCounter = 0
	
			If (Err.Number <> 0) Then
				sErrorDescription = "Error getting the event log records." & vbCrLf & FormatError(Err.Description)
			End If

			If Len(sErrorDescription) = 0 Then
				If Not (rsEmailDetails.bof And rsEmailDetails.eof) Then
					Do Until rsEmailDetails.eof
			
						If iLastEventID <> rsEmailDetails.Fields("ID").value Then
					
							EventCounter = EventCounter + 1
							Response.Write(CStr(EventCounter))

							sEmailInfo = sEmailInfo & StrDup(Len(rsEmailDetails.Fields("Name").Value) + 30, "-") & vbCrLf
							sEmailInfo = sEmailInfo & "Event Name : " & rsEmailDetails.Fields("Name").Value & vbCrLf
							sEmailInfo = sEmailInfo & StrDup(Len(rsEmailDetails.Fields("Name").Value) + 30, "-") & vbCrLf
					
							sEmailInfo = sEmailInfo & "Mode :		" & rsEmailDetails.Fields("Mode").Value & vbCrLf & vbCrLf
					
							sEmailInfo = sEmailInfo & "Start Time :	" & ConvertSQLDateToLocale(rsEmailDetails.Fields("DateTime").Value) & " " & ConvertSqlDateToTime(rsEmailDetails.Fields("DateTime").Value) & vbCrLf
							If IsDBNull(rsEmailDetails.Fields("EndTime").Value) Then
								sEmailInfo = sEmailInfo & "End Time :	" & vbCrLf
							Else
								sEmailInfo = sEmailInfo & "End Time :	" & ConvertSQLDateToLocale(rsEmailDetails.Fields("DateTime").Value) & " " & ConvertSqlDateToTime(rsEmailDetails.Fields("EndTime").Value) & vbCrLf
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
				
						sEmailInfo = sEmailInfo & vbCrLf & vbCrLf & vbCrLf
				
						rsEmailDetails.Movenext()
					Loop
		
					Response.Write("<INPUT type=hidden name=txtEventDeleted id=txtEventDeleted value=0>" & vbCrLf)
			
				Else
					Response.Write("<INPUT type=hidden name=txtEventDeleted id=txtEventDeleted value=1>" & vbCrLf)
				End If
			End If
	
			rsEmailDetails.close()
			rsEmailDetails = Nothing

			Response.Write("<INPUT type=hidden name=txtBody id=txtBody value=""" & Replace(sEmailInfo, """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type=hidden name=txtSubject id=txtSubject value=""" & Replace(cmdEmailDetails.Parameters("subject").Value, """", "&quot;") & """>" & vbCrLf)
	
			cmdEmailDetails = Nothing
			objUtilities = Nothing
	
		%>
	</form>



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

	<script type="text/javascript" id="scptGeneralFunctions">
<!--

	function populateEmailList() {
		var sAddLine = '';
	
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
