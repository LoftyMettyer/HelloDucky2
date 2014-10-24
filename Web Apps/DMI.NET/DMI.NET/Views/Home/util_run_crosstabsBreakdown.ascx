<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<form action="util_run_CrossTabsBreakdown" method="post" id="frmBreakdown" name="frmBreakdown">
		<input type="hidden" id="txtMode" name="txtMode" value="<%=Session("CT_Mode")%>">
		<input type="hidden" id="txtHor" name="txtHor" value="<%=Session("CT_Hor")%>">
		<input type="hidden" id="txtVer" name="txtVer" value="<%=Session("CT_Ver")%>">
		<input type="hidden" id="txtPgb" name="txtPgb" value="<%=Session("CT_Pgb")%>">
		<input type="hidden" id="txtIntersectionType" name="txtIntersectionType" value=0>
		<input type="hidden" id="txtCellValue" name="txtCellValue" value="<%=Session("CT_CellValue")%>">
</form>

<%
	Dim objCrossTab = CType(Session("objCrossTab" & Session("CT_UtilID")), CrossTab)
	
		If Session("CT_Mode") = "BREAKDOWN" Then

		If objCrossTab.CrossTabType = CrossTabType.cttAbsenceBreakdown Then
			Response.Write("<label style='font-weight: bold;'>Absence Breakdown Cell Breakdown</label>")
		ElseIf objCrossTab.CrossTabType = CrossTabType.ctt9GridBox Then
			Response.Write("<label id='nineBoxGridCellBreakdownLabel' style='font-weight: bold;'></label>")
		Else
			Response.Write("<label style='font-weight: bold;'>Cross Tabs Cell Breakdown</label>")
		End If
		
%>

		<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%">
				<tr height="1">
						<td>
								<table height="100%" width="100%" class="invisible" cellspacing="0" cellpadding="0">

										<%
			
												If objCrossTab.PageBreakColumnName <> "<None>" Then
														Response.Write("<TR HEIGHT=5>" & vbCrLf)
														Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
														Response.Write("  <TD WIDTH=""50%"">" & objCrossTab.PageBreakColumnName & " :</TD>" & vbCrLf)
														Response.Write("  <TD WIDTH=""50%""><INPUT id=txtPgb class=""text textdisabled"" name=txtPgb value="" " & _
																						objCrossTab.ColumnHeading(2, Session("CT_Pgb")) & _
																						""" style=""WIDTH: 100%"" disabled=""disabled""></TD>" & vbCrLf)
														Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
														Response.Write("</TR>" & vbCrLf)
														Response.Write("<TR HEIGHT=5><TD></TD></TR>" & vbCrLf)
												End If

												Response.Write("<TR HEIGHT=5>" & vbCrLf)
												Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
												Response.Write("  <TD WIDTH=""50%"">" & objCrossTab.HorizontalColumnName & " :</TD>" & vbCrLf)
												Response.Write("  <TD WIDTH=""50%""><INPUT id=txtHor name=txtHor value="" ")

												If CLng(Session("CT_Hor")) > CLng(objCrossTab.ColumnHeadingUbound(0)) Then
														Response.Write("<All>")
												Else
														Response.Write(RoundValuesInRange(objCrossTab.ColumnHeading(0, Session("CT_Hor"))))
												End If
												Response.Write(""" style=""WIDTH: 100%"" class=""text textdisabled"" disabled=""disabled""></TD>" & vbCrLf)
												Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
												Response.Write("</TR>" & vbCrLf)
												Response.Write("<TR HEIGHT=5><TD></TD></TR>" & vbCrLf)

												Response.Write("<TR HEIGHT=5>" & vbCrLf)
												Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
												Response.Write("  <TD>" & objCrossTab.VerticalColumnName & " :</TD>" & vbCrLf)
												Response.Write("  <TD><INPUT id=txtVer name=txtVer value="" ")
												If CLng(Session("CT_Ver")) > CLng(objCrossTab.ColumnHeadingUbound(1)) Then
														Response.Write("<All>")
												Else
														Response.Write(RoundValuesInRange(objCrossTab.ColumnHeading(1, Session("CT_Ver"))))
												End If
												Response.Write(""" style=""WIDTH: 100%"" class=""text textdisabled"" disabled=""disabled""></TD>" & vbCrLf)
												Response.Write("</TR>" & vbCrLf)
												Response.Write("<TR HEIGHT=5><TD></TD></TR>" & vbCrLf)


												Response.Write("<TR HEIGHT=5>" & vbCrLf)
												Response.Write("  <TD WIDTH=50>&nbsp;</TD>" & vbCrLf)
											Response.Write("  <TD>" & objCrossTab.IntersectionTypeValue(Session("CT_IntersectionType")) & " :</TD>" & vbCrLf)
												Response.Write("  <TD><INPUT id=txtCellValue name=txtCellValue value="" ")
				
											If objCrossTab.CrossTabType = CrossTabType.cttAbsenceBreakdown Then
												Response.Write(objCrossTab.OutputArrayDataUBound)
											Else
												Response.Write(Session("CT_CellValue"))
											End If

												Response.Write(""" style=""WIDTH: 100%"" class=""text textdisabled"" disabled=""disabled""></TD>" & vbCrLf)
												Response.Write("</TR>" & vbCrLf)
												Response.Write("<TR HEIGHT=5><TD></TD></TR>" & vbCrLf)

										%>
										<tr>
												<td></td>
												<td colspan="2">
													<table id="ssOutputBreakdown"></table>													
												</td>
										</tr>
								</table>
						</td>
				</tr>
		</table>	
		<%

			Response.Write("<script type=""text/javascript"">" & vbCrLf)			
			Response.Write("function util_run_crosstabsBreakdown_window_onload()" & vbCrLf)
			Response.Write("{" & vbCrLf & vbCrLf)

			Response.Write("    var colNames = [];" & vbCrLf)
			Response.Write("    var colData = [];" & vbCrLf)
			Response.Write("    var colMode = [];" & vbCrLf)
			Response.Write("    var value;" & vbCrLf)
			Response.Write("    var i;" & vbCrLf)
			Response.Write("    var sColumnName;" & vbCrLf)
			Response.Write("    var iCount2;" & vbCrLf)
			Response.Write("    var obj;" & vbCrLf)
		
		Response.Write("  colNames.push('" & CleanStringForJavaScript(objCrossTab.BaseTableName) & "');" & vbCrLf)
		Response.Write("	colMode.push({ name: '" & CleanStringForJavaScript(objCrossTab.BaseTableName) & "' });" & vbCrLf)	
		
		' Absence Breakdown
			If objCrossTab.CrossTabType = CrossTabType.cttAbsenceBreakdown Then

				Response.Write("  colNames.push('Start Date');" & vbCrLf)
				Response.Write("	colMode.push({ name: 'Start Date' });" & vbCrLf)

				Response.Write("  colNames.push('End Date');" & vbCrLf)
				Response.Write("	colMode.push({ name: 'End Date' });" & vbCrLf)

				Response.Write("  colNames.push(""" & CleanStringForJavaScript(objCrossTab.ColumnHeading(0, Session("CT_Hor"))) & "'s taken" & """);" & vbCrLf)
				Response.Write("	colMode.push({ name: """ & CleanStringForJavaScript(objCrossTab.ColumnHeading(0, Session("CT_Hor"))) & "'s taken" & """ });" & vbCrLf)
				
			End If
		
		If objCrossTab.IntersectionColumn = True Then
			Response.Write("  colNames.push('" & CleanStringForJavaScript(objCrossTab.IntersectionColumnName) & "');" & vbCrLf)
			Response.Write("	colMode.push({ name: '" & CleanStringForJavaScript(objCrossTab.IntersectionColumnName) & "' });" & vbCrLf)
		End If

		
			Dim objData As String()
			For intCount = 1 To objCrossTab.OutputArrayDataUBound
								
				Response.Write("  obj = {};" & vbCrLf)
				objData = Split(objCrossTab.OutputArrayData(intCount), vbTab)
				For intCount2 = 0 To UBound(objData)
					Response.Write("  obj[colNames[" & intCount2 & "]] = '" & objData(intCount2) & "';" & vbCrLf)
				Next
				Response.Write("  colData.push(obj);")
			Next
		
			Response.Write("	$('#ssOutputBreakdown').jqGrid({data: colData, datatype: 'local', colNames: colNames, height: 400, colModel: colMode, autowidth: true" & vbCrLf)
			Response.Write("   ,rowNum:1000000")
			Response.Write("	, cmTemplate: { editable: true }});" & vbCrLf & vbCrLf)

			Response.Write("$('#nineBoxGridCellBreakdownLabel').text('9-Box Grid Cell Breakdown: ' + $('#nineBoxR' + " & Session("CT_Ver") - 1 & " + 'C' + " & Session("CT_Hor") - 1 & ").attr('data-titlevalue'));" & vbCrLf)
			
		Response.Write("}" & vbCrLf)
			Response.Write("</script>" & vbCrLf & vbCrLf)

			objCrossTab = Nothing

		%>


<script type="text/javascript">

	$("#reportbreakdownframe").attr("data-framesource", "UTIL_RUN_CROSSTABSBREAKDOWN");
	
	$("#reportframe").show();
	$("#reportdataframe").hide();
	$("#reportworkframe").hide();
	$("#reportbreakdownframe").show();

	$('.ui-dialog-buttonpane #cmdClose').hide();
	<%If Session("SSIMode") = False Then%>
	$('.ui-dialog-buttonpane #cmdCancel').button('enable');
	$('.ui-dialog-buttonpane #cmdClose').show();
	<%Else%>
	$('#divReportButtons #cmdCancel').button('enable');
	<%End If%>
	setTimeout("util_run_crosstabsBreakdown_window_onload()", 100);

</script>

<%
		End If
		%>
