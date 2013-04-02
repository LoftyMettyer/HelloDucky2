<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<object
    classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
    id="Microsoft_Licensed_Class_Manager_1_0">
    <param name="LPKPath" value="lpks/main.lpk">
</object>

	
<script type="text/javascript">

    function populateKey()
    {
        var strKey, strDescription, strCode;
        var lngColour;
        var strControlName = '';
	
        frmKey.ctlKey.Clear_Key();
	
        for (var i=1; i<=frmKeyInfo.key_Count.value; i++)
        {
            strControlName = 'key_ID'+i;
            strKey = document.getElementById(strControlName).getAttribute('value');
            strControlName = 'key_Name'+i;
            strDescription = document.getElementById(strControlName).getAttribute('value');
            strControlName = 'key_Code'+i;
            strCode = document.getElementById(strControlName).getAttribute('value');
            strControlName = 'key_Colour'+i;
            lngColour = Number(document.getElementById(strControlName).getAttribute('value'));
		
            frmKey.ctlKey.Add_Key(strKey, strDescription, strCode, lngColour);
        }
	
        frmKey.ctlKey.Sort();
	
        if (frmKeyInfo.txtHasMultiple.value == '1')
        {
            strKey = "EVENT_MULTIPLE";
            strDescription = "Multiple Events";
            strCode = ".";
            lngColour = 16777215;
	
            frmKey.ctlKey.Add_Key(strKey, strDescription, strCode, lngColour);
        }
	
        return true;
    }

</script>

<form id="frmKey" name="frmKey">
    <table align="center" class="outline" cellpadding="0" cellspacing="0" width="100%" height="100%">
        <tr>
            <td>
                <table class="invisible" cellspacing="0" cellpadding="2" width="100%" height="100%">
                    <tr height="5" valign="top">
                        <td width="5"></td>
                        <td height="100%" width="100%" align="left">Key :
                            <br>
                            <object
                                classid="CLSID:8E2F1EF1-3812-4678-A084-16384DE3EA6D"
                                codebase="cabs/COAInt_CalRepKey.cab#version=1,0,0,2"
                                id="ctlKey"
                                name="ctlKey"
                                width="100%"
                                height="95"
                                style="width: 100%; height: 95px">
                            </object>
                        </td>
                        <td width="5"></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</form>

	<form id=frmKeyInfo name=frmKeyInfo style="visibility:hidden;display:none">
<%
    Dim rsColours As Object
    Dim intColourCount As Integer
    Dim intNextIndex As Integer
    Dim mavAvailableColours(,)
    Dim cmdColours As Object
	
  intColourCount = 0
  intNextIndex = 0
  ReDim mavAvailableColours(3, intNextIndex)
  
    cmdColours = CreateObject("ADODB.Command")
	cmdColours.CommandText = "spASRIntGetCalendarColours"
	cmdColours.CommandType = 4
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
    Dim objCalendar As Object
  
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
    Response.Write("<INPUT name=key_Count id=key_Count value=" & intLegendCount & ">" & vbCrLf)
	
    If objCalendar.HasMultipleEvents Then
        Response.Write("<INPUT name=txtHasMultiple id=txtHasMultiple value='1'>" & vbCrLf)
    Else
        Response.Write("<INPUT name=txtHasMultiple id=txtHasMultiple value='0'>" & vbCrLf)
    End If

%>
		<INPUT type="hidden" id=txtCalRep_UtilID name=txtCalRep_UtilID value=<%Session("CalRepUtilID").ToString()%>>
	</form>


<%
    objCalendar = Nothing
%>
    
<script type="text/javascript">
    populateKey();
</script>
