<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="ADODB" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/bundles/utilities_calendarreport_run")%>" type="text/javascript"></script>  

	<link href="<%= Url.LatestContent("~/Content/calendarreport.css")%>" rel="stylesheet" type="text/css" />
	
Key :
<table id="tblCalendarReportKey">
</table>


	<form id=frmKeyInfo name=frmKeyInfo style="visibility:hidden;display:none">
<%
	Dim rsColours As Recordset
	Dim intColourCount As Integer
	Dim intNextIndex As Integer
	Dim mavAvailableColours(,)
	Dim cmdColours As Command
	
  intColourCount = 0
  intNextIndex = 0
  ReDim mavAvailableColours(3, intNextIndex)
  
	cmdColours = New Command
	cmdColours.CommandText = "spASRIntGetCalendarColours"
	cmdColours.CommandType = CommandTypeEnum.adCmdStoredProc
	cmdColours.ActiveConnection = Session("databaseConnection")

    Err.Clear()
    rsColours = cmdColours.Execute
  
  With rsColours
    If not (.BOF And .EOF) Then
			.MoveFirst
			Do While Not .EOF
			  ReDim Preserve mavAvailableColours(3, intNextIndex)
			  
			  mavAvailableColours(0, intNextIndex) = .Fields("ColValue").Value 
			  mavAvailableColours(1, intNextIndex) = Hex(.Fields("ColValue").Value)
			  mavAvailableColours(2, intNextIndex) = .Fields("ColDesc").Value 
			  mavAvailableColours(3, intNextIndex) = .Fields("WordColourIndex").Value 

			  intNextIndex = UBound(mavAvailableColours, 2) + 1
			  
			  .MoveNext
			Loop
		end if
  End With
  rsColours.Close
    rsColours = Nothing
    cmdColours = Nothing
	
    Dim rsKey As Object
    Dim objCalendar As HR.Intranet.Server.CalendarReport
  
    Dim intCount As Integer
    Dim strEventID As String
    Dim blnNewEvent As Boolean
    Dim intColourIndex As Integer
    Dim intColourMax As Integer
	Dim intLegendCount As Integer
    Dim intNewIndex As Integer
	   
    Dim mavLegend(,)
	
    
    objCalendar = Session("objCalendar" & Session("CalRepUtilID"))
	
    rsKey = objCalendar.EventsRecordset

  strEventID = vbNullString

  ReDim mavLegend(3, 0)
  
    intLegendCount = 0
  
  intColourMax = UBound(mavAvailableColours, 2)
	
	With rsKey
		If Not (.BOF And .EOF) Then
		    
		  .MoveFirst
		  Do While Not .EOF
				
        If strEventID <> .Fields(objCalendar.EventIDColumn).Value Then
          strEventID = .Fields(objCalendar.EventIDColumn).Value
          
          blnNewEvent = True
          For intCount = 1 To UBound(mavLegend, 2) Step 1
            If mavLegend(0, intCount) = strEventID Then
              blnNewEvent = False
            End If
          next
          
          If blnNewEvent Then
            intNewIndex = UBound(mavLegend, 2) + 1
            
            ReDim Preserve mavLegend(3, intNewIndex)
            mavLegend(0, intNewIndex) = strEventID
            mavLegend(1, intNewIndex) = Left(.Fields("Name").Value, 50)
            mavLegend(2, intNewIndex) = Left(.Fields("Legend").Value, 2)
          
            intColourIndex = (intNewIndex - 1) Mod intColourMax
            mavLegend(3, intNewIndex) = mavAvailableColours(0, intColourIndex)
            
                        Response.Write("<INPUT name=key_ID" & intNewIndex & " id=key_ID" & intNewIndex & " value=""" & mavLegend(0, intNewIndex) & """>" & vbCrLf)
                        Response.Write("<INPUT name=key_Name" & intNewIndex & " id=key_Name" & intNewIndex & " value=""" & mavLegend(1, intNewIndex) & """>" & vbCrLf)
                        Response.Write("<INPUT name=key_Code" & intNewIndex & " id=key_Code" & intNewIndex & " value=""" & mavLegend(2, intNewIndex) & """>" & vbCrLf)
                        Response.Write("<INPUT name=key_Colour" & intNewIndex & " id=key_Colour" & intNewIndex & " value=""" & mavLegend(3, intNewIndex) & """>" & vbCrLf)
                    End If
        End If
				
				.MoveNext
			loop
		end if
	end with 

	intLegendCount = UBound(mavLegend, 2)
	Response.Write("<input name=key_Count id=key_Count value=" & intLegendCount & ">" & vbCrLf)
	
	If objCalendar.HasMultipleEvents Then
		Response.Write("<input name=txtHasMultiple id=txtHasMultiple value='1'>" & vbCrLf)
	Else
		Response.Write("<input name=txtHasMultiple id=txtHasMultiple value='0'>" & vbCrLf)
	End If

	objCalendar = Nothing
%>
		<input type="hidden" id="txtCalRep_UtilID" name="txtCalRep_UtilID" value='<%Session("CalRepUtilID").ToString()%>'>
	</form>
 
<script type="text/javascript">
    populateKey();
</script>
