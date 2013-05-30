<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<script src="<%: Url.Content("~/bundles/utilities_calendarreport_run")%>" type="text/javascript"></script>  



<form name="frmCalendar" id="frmCalendar">
<table align=center class="invisible" cellPadding=0 cellSpacing=0 width=100%>
	<tr height=30>
		<TD align=right nowrap width=100%>
			<table class="invisible" cellspacing="0" cellpadding="0" width=100% height=100%>
<%
	Session("CalendarFrameLoaded") = False
	
    Dim rsCalendarBaseInfo As Object
    Dim intBaseRecordCount As Integer
    Dim blnNewBaseRecord As Boolean
    Dim strTempRecordDesc As String
    Dim strBaseRecDesc As String
    Dim strConvertedBaseRecDesc As String
    Dim intDescEmpty As Integer
    Dim blnDescEmpty As Boolean
    Dim objCalendar As HR.Intranet.Server.CalendarReport
    Dim intBaseRecCount As Integer
    Dim lngCurrentRecordID As Long
    Dim strCurrentBaseRegion As String
	
    objCalendar = Session("objCalendar" & Session("CalRepUtilID"))
	' Create the reference to the DLL (Report Class)
	
    rsCalendarBaseInfo = objCalendar.BaseRecordset

	intBaseRecCount = 0 
	strBaseRecDesc = vbNullString 
	
	if not (rsCalendarBaseInfo.bof and rsCalendarBaseInfo.eof) then
		rsCalendarBaseInfo.movefirst
		do until rsCalendarBaseInfo.eof
				
			strTempRecordDesc = objCalendar.ConvertDescription(rsCalendarBaseInfo.Fields("Description1").Value,rsCalendarBaseInfo.Fields("Description2").Value,rsCalendarBaseInfo.Fields("DescriptionExpr").Value)

	    blnDescEmpty = (strTempRecordDesc = "")
	    If blnDescEmpty Then
	      intDescEmpty = intDescEmpty + 1
	    Else
	      intDescEmpty = 0
	    End If
		    
	    if objCalendar.GroupByDescription then
				If (strTempRecordDesc <> strBaseRecDesc) Or (blnDescEmpty And Int(intDescEmpty = 1)) Then
	        blnNewBaseRecord = True
	        blnDescEmpty = False
		      
		      strBaseRecDesc = strTempRecordDesc
	        strConvertedBaseRecDesc = strTempRecordDesc
		        
	        If Len(Trim(objCalendar.StaticRegionColumn)) > 0 Then
	          strCurrentBaseRegion = rsCalendarBaseInfo.Fields("Region").Value
	        end If
	        intBaseRecordCount = intBaseRecordCount + 1
	      end if
	      lngCurrentRecordID = rsCalendarBaseInfo.Fields(objCalendar.BaseIDColumn).Value
		      				
			else
	      If rsCalendarBaseInfo.Fields(objCalendar.BaseIDColumn).Value <> lngCurrentRecordID Then
	        blnNewBaseRecord = True
		        
	        strConvertedBaseRecDesc = strTempRecordDesc
		       
	        lngCurrentRecordID = rsCalendarBaseInfo.Fields(objCalendar.BaseIDColumn).Value
	        If Len(Trim(objCalendar.StaticRegionColumn)) > 0 Then
						strCurrentBaseRegion = rsCalendarBaseInfo.Fields("Region").Value
	        End If
	        intBaseRecordCount = intBaseRecordCount + 1
	      End If

			end if

			if blnNewBaseRecord then
				intBaseRecCount = intBaseRecCount + 1
				 
                Response.Write("<tr>" & vbCrLf)
                Response.Write("<td width=""100%"">" & vbCrLf)
                Response.Write("<object classid=""CLSID:252D73AF-D7C6-4833-8539-A2C0293950B1""" & vbCrLf)
                Response.Write("				CODEBASE=""cabs/COAInt_CalRepRecord.cab#version=1,0,0,2""" & vbCrLf)
                Response.Write("				id=ctlCalRec_" & intBaseRecCount & vbCrLf)
                Response.Write("				name=ctlCalRec_" & intBaseRecCount & vbCrLf)
                Response.Write("				style=""VISIBILITY: visible; WIDTH: 100%""" & vbCrLf)
                Response.Write("				width=""100%"">" & vbCrLf)
                Response.Write("					<PARAM NAME=""BaseDesc"" VALUE=""" & Replace(Replace(strConvertedBaseRecDesc, "&", "&&"), """", "&quot;") & """>" & vbCrLf)
				
                If objCalendar.GroupByDescription Then
                    Response.Write("					<PARAM NAME=""BaseDescTag"" VALUE=""-1"">" & vbCrLf)
                    Response.Write("					<PARAM NAME=""Region"" VALUE="""">" & vbCrLf)
                Else
                    Response.Write("					<PARAM NAME=""BaseDescTag"" VALUE=""" & lngCurrentRecordID & """>" & vbCrLf)
                    Response.Write("					<PARAM NAME=""Region"" VALUE=""" & strCurrentBaseRegion & """>" & vbCrLf)
                End If
                Response.Write("</object>" & vbCrLf)
				
                Response.Write("<script type=""text/javascript"">")
                Response.Write("function ctlCalRec_CalDateClick" & intBaseRecCount & "(pvarLabel) {")

                Response.Write("    var strKey;" & vbCrLf)
                Response.Write("	var lngOriginalLeft;" & vbCrLf)
                Response.Write("	var lngOriginalTop;" & vbCrLf)
                Response.Write("	var strDate;" & vbCrLf)
                Response.Write("	var strSession = new String('');" & vbCrLf)
                Response.Write("	var sURL;" & vbCrLf)
				
                Response.Write("	var	CALDATES_BOXWIDTH = new Number(200);" & vbCrLf)
                Response.Write("	var	CALDATES_BOXHEIGHT = new Number(200);" & vbCrLf)

                Response.Write("    var CalRec = $(""#ctlCalRec_" & intBaseRecCount & """)[0];" & vbCrLf & vbCrLf)
                
                Response.Write("	with (frmEventDetails)" & vbCrLf)
                Response.Write("		{" & vbCrLf)
                Response.Write("		if (Number(CalRec.TagInfo_Get(pvarLabel.Tag, ""HAS_EVENT"")) > 0)" & vbCrLf)
                Response.Write("			{" & vbCrLf)
                Response.Write("			strDate = CalRec.ConvertTagDateToString(CalRec.TagInfo_Get(pvarLabel.Tag, ""DATE""));" & vbCrLf)
                Response.Write("			strSession = CalRec.TagInfo_Get(pvarLabel.Tag, ""SESSION"");" & vbCrLf)
			
                Response.Write("			txtBaseIndex.value = " & intBaseRecCount & ";" & vbCrLf)
                Response.Write("			txtLabelIndex.value = pvarLabel.Index;" & vbCrLf)
				
                If objCalendar.IncludeBankHolidays_Enabled Then
                    Response.Write("			txtShowRegion.value = 1;" & vbCrLf)
                Else
                    Response.Write("			txtShowRegion.value = 0;" & vbCrLf)
                End If
				
                If objCalendar.IncludeWorkingDaysOnly_Enabled Then
                    Response.Write("			txtShowWorkingPattern.value = 1;" & vbCrLf)
                Else
                    Response.Write("			txtShowWorkingPattern.value = 0;" & vbCrLf)
                End If
				
                Response.Write("			txtBreakdownCaption.value = 'Calendar Report Breakdown - ' + strDate + ' ' + strSession.toLowerCase();" & vbCrLf)
				
                Response.Write("			sURL = ""util_run_calendarreport_breakdown"" +" & vbCrLf & _
                    """?txtBreakdownCaption="" + escape(frmEventDetails.txtBreakdownCaption.value) +" & vbCrLf & _
                    """&txtShowRegion="" + escape(frmEventDetails.txtShowRegion.value) + " & vbCrLf & _
                    """&txtShowWorkingPattern="" + escape(frmEventDetails.txtShowWorkingPattern.value) +" & vbCrLf & _
                    """&txtBaseIndex="" + escape(frmEventDetails.txtBaseIndex.value) +" & vbCrLf & _
                    """&CalRepUtilID="" + escape(frmCalendar.txtCalRep_UtilID.value) +" & vbCrLf & _
                    """&txtLabelIndex="" + escape(frmEventDetails.txtLabelIndex.value);" & vbCrLf & _
                    "openDialog(sURL, 370,475, ""yes"", ""no"");" & vbCrLf)

                Response.Write("			}" & vbCrLf)
                Response.Write("		}	" & vbCrLf)
                Response.Write("	}	" & vbCrLf)

                Response.Write("OpenHR.addActiveXHandler(""ctlCalRec_" & intBaseRecCount & """, ""CalDateClick"", ctlCalRec_CalDateClick" & intBaseRecCount & ");" & vbCrLf)

                Response.Write("</script>")

                Response.Write("</td>" & vbCrLf)
                Response.Write("</tr>" & vbCrLf)
            End If
			
            objCalendar.BaseIndex_Add(CInt(intBaseRecordCount), CLng(lngCurrentRecordID))
			
			blnNewBaseRecord = False
							
			rsCalendarBaseInfo.movenext
		loop
	else
	
	end if
	
	Session("BaseControlCount") = intBaseRecordCount
	
	if objCalendar.GroupByDescription then
		Session("GroupByDesc") = 1
	else
		Session("GroupByDesc") = 0
	end if

    objCalendar = Nothing
	
%>
			</table>
		</td>
	</tr>
</table>

    <input type="hidden" name="txtBaseCtlCount" id="txtBaseCtlCount" value="<%=Session("BaseControlCount")%>">
    <input type="hidden" name="txtGroupByDesc" id="txtGroupByDesc" value="<%=Session("GroupByDesc")%>">
    <input type="hidden" id="txtCalRep_UtilID" name="txtCalRep_UtilID" value='<%=Session("CalRepUtilID").ToString()%>'>
</form>

<form id="frmEventDetails" name="frmEventDetails" target="eventDetails" action="util_run_calendarreport_breakdown" method="post" style="visibility: hidden; display: none">
    <input type="hidden" name="txtBreakdownCaption" id="txtBreakdownCaption">
    <input type="hidden" name="txtShowRegion" id="txtShowRegion">
    <input type="hidden" name="txtShowWorkingPattern" id="txtShowWorkingPattern">
    <input type="hidden" name="txtBaseIndex" id="txtBaseIndex">
    <input type="hidden" name="txtLabelIndex" id="txtLabelIndex">
</form>

<script type="text/javascript">
    util_run_calendarreport_calendar_window_onload();
</script>
