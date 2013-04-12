<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>

<link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />
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

<%
    ' Write the prompted values from the calling form into a session variable.
    Dim j As Integer
    Dim aPrompts(1, 0) As String
    Dim sKey As String
    j = 0
    ReDim Preserve aPrompts(1, 0)
    For i = 0 To (Request.Form.Count) - 1
        sKey = Request.Form.Keys(i)
        If ((UCase(Left(sKey, 7)) = "PROMPT_") And (Mid(sKey, 8, 1) <> "3")) Or _
            (UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
            ReDim Preserve aPrompts(1, j)
            If (UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
                aPrompts(0, j) = "prompt_3_" & Mid(sKey, 11)
                aPrompts(1, j) = UCase(Request.Form.Item(i))
            Else
                aPrompts(0, j) = sKey
                Select Case Mid(sKey, 8, 1)
                    Case "2"
                        ' Numeric. Replace locale decimal point with '.'
                        aPrompts(1, j) = Replace(Request.Form.Item(i), Session("LocaleDecimalSeparator"), ".")
                    Case "4"
                        ' Date. Reformat to match SQL's mm/dd/yyyy format.
                        aPrompts(1, j) = convertLocaleDateToSQL(Request.Form.Item(i))
                    Case Else
                        aPrompts(1, j) = Request.Form.Item(i)
                End Select
            End If
            j = j + 1
        End If
    Next
    Session("TestPrompts") = aPrompts
%>

<html>
<head runat="server">
    <title>OpenHR Intranet</title>
    
    <script ID="clientEventHandlersJS"  type="text/javascript">
        function util_test_expression_onload() {

            if (txtDisplay.value != "False") {

                // Hide the 'please wait' message.
                trPleaseWait1.style.visibility='hidden';
                trPleaseWait1.style.display='none';
                trPleaseWait2.style.visibility='hidden';
                trPleaseWait2.style.display='none';
                trPleaseWait3.style.visibility='hidden';
                trPleaseWait3.style.display='none';
                trPleaseWait4.style.visibility='hidden';
                trPleaseWait4.style.display='none';
                trPleaseWait5.style.visibility='hidden';
                trPleaseWait5.style.display='none';

                var bdyMain = $("#bdyMain");

                // Resize the grid to show all prompted values.
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
            else {
                self.close();
            }
        }

    </script>


</head>

<body id="bdyMain">    
    <div data-framesource="util_test_expression">

    
    <table align=center class="outline" cellPadding=5 cellSpacing=0>
	<TR>
		<TD>
			<table class="invisible" cellspacing="0" cellpadding="0">
				<tr> 
			    <td colspan=5 height=10></td>
			  </tr>

			  <tr id=trPleaseWait1> 
					<td width=20></td>
			    <td align=center colspan=3> 
<%
	if Request.form("type") = 11 then
        Response.Write("Testing Filter")
	else
		if Request.form("type") = 12 then
            Response.Write("Testing Calculation")
		else
            Response.Write("Testing Expression")
		end if
	end if
%>						
			    </td>
					<td width=20></td>
			  </tr>

			  <tr id=trPleaseWait4 height=10> 
					<td colspan=5></td>
			  </tr>

			  <tr id=trPleaseWait2> 
					<td width=20></td>
			    <td align=center colspan=3> 
						Please Wait...
			    </td>
					<td width=20></td>
			  </tr>

			  <tr id=trPleaseWait5 height=20> 
					<td colspan=5></td>
			  </tr>

			  <tr id=trPleaseWait3> 
					<td width=20></td>
			    <td align=center colspan=3> 
						<INPUT id=Cancel name=Cancel class="btn" type=button value=OK style="WIDTH: 80px" width="80" 
						    onclick="self.close()"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
			    </td>
					<td width=20></td>
			  </tr>


<%
	dim fOK
	dim fDisplay
	dim sUtilType
    Dim sUtilType2
    Dim objExpression
    Dim iExprType As Integer
    Dim iReturnType As Integer
    Dim iValidityCode As Integer
    Dim sValidityMessage As String
    Dim sFilterCode As String
    Dim iRecCount As Integer
    Dim sMsg1 As String
    Dim sMsg As String

	dim mvarUDFsRequired()
	redim preserve mvarUDFsRequired(0)

	fOK = true
	fDisplay = false
	
	if Request.form("type") = "11" then
		sUtilType = "Filter"
		sUtilType2 = "filter"
	else
		sUtilType = "Calculation"
		sUtilType2 = "calculation"
	end if
		
	' Get the server DLL to test the expression definition
    objExpression = New HR.Intranet.Server.Expression

	' Pass required info to the DLL
    objExpression.Username = Session("username")
    CallByName(objExpression, "Connection", CallType.Let, Session("databaseConnection"))
    
	if fok then 
		if Request.form("type") = 11 then
			iExprType = 11
			iReturnType = 3
		else
			iExprType = 10
			iReturnType = 0
		end if
				
        fOK = objExpression.Initialise(CLng(Request.Form("tableID")), 0, CInt(iExprType), CInt(iReturnType))
	end if

	if fok then 
        fOK = objExpression.SetExpressionDefinition(CStr(Request.Form("components1")), "", "", "", "", "")
	end if

	if fok then 
	  iValidityCode = objExpression.ValidateExpression

	  If iValidityCode > 0 Then
			fDisplay = true
            Response.Write("			  <tr>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
            Response.Write("						<H3>Error Testing " & sUtilType & "</H3>" & vbCrLf)
            Response.Write("			    </td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			  </tr>" & vbCrLf)
            Response.Write("			  <tr>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=center colspan=3> " & vbCrLf)

			sValidityMessage = objExpression.ValidityMessage(cint(iValidityCode))
			sValidityMessage = replace(sValidityMessage, vbcr, "<BR>") 

            Response.Write("						 " & sValidityMessage & vbCrLf)

            Response.Write("			    </td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			  </tr>" & vbCrLf)
            Response.Write("			  <tr>" & vbCrLf)
            Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
            Response.Write("			  </tr>" & vbCrLf)
            Response.Write("			  <tr> " & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
%>			
									<INPUT id=Button1 name=Cancel type=button class="btn" value=OK style="WIDTH: 80px" width=80 
									    onclick="self.close()"
	                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
	                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
	                                    onfocus="try{button_onFocus(this);}catch(e){}"
	                                    onblur="try{button_onBlur(this);}catch(e){}" />
<%
    Response.Write("			    </td>" & vbCrLf)
    Response.Write("					<td width=20></td>" & vbCrLf)
    Response.Write("			  </tr>" & vbCrLf)
		end if
	end if

	if fok and (fDisplay = false) then 
		objExpression.SetPromptedValues(aPrompts)

		sFilterCode = objExpression.RuntimeFilterCode

		' Create dynamic User defined functions
objExpression.UDFFilterCode(True)
		
		iRecCount = 0
		
		fDisplay = true
		if len(sFilterCode) = 0 then
			sMsg1 = "Testing " & sUtilType
			sMsg = "Your " & sUtilType2 & " is defined correctly." 
		else
			iRecCount = objExpression.TestFilterCode(cstr(sFilterCode))
			
			if iRecCount < 0 then
				sMsg1 = "Error Testing " & sUtilType
				sMsg = "Error running the test " & sUtilType2 & " SQL code."
			else
				sMsg1 = "Testing " & sUtilType
				sMsg = "Your " & sUtilType2 & " is defined correctly.<BR><BR>" & _
					"You have permission to view " & iRecCount & " record" 
					
				if(iRecCount <> 1) then
					sMsg = sMsg & "s"
				end if
					sMsg = sMsg & " using this filter." & vbcrlf
			end if
		end if

		' Create dynamic User defined functions
objExpression.UDFFilterCode(False)

				
Response.Write("			  <tr>" & vbCrLf)
Response.Write("					<td width=20></td>" & vbCrLf)
Response.Write("			    <td align=center colspan=3> " & vbCrLf)
Response.Write("						<H3>" & sMsg1 & "</H3>" & vbCrLf)
Response.Write("			    </td>" & vbCrLf)
Response.Write("					<td width=20></td>" & vbCrLf)
Response.Write("			  </tr>" & vbCrLf)
Response.Write("			  <tr>" & vbCrLf)
Response.Write("					<td width=20></td>" & vbCrLf)
Response.Write("			    <td align=center colspan=3> " & vbCrLf)
Response.Write("						 " & sMsg & vbCrLf)
Response.Write("			    </td>" & vbCrLf)
Response.Write("					<td width=20></td>" & vbCrLf)
Response.Write("			  </tr>" & vbCrLf)
Response.Write("			  <tr>" & vbCrLf)
Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
Response.Write("			  </tr>" & vbCrLf)
Response.Write("			  <tr> " & vbCrLf)
Response.Write("					<td width=20></td>" & vbCrLf)
Response.Write("			    <td align=center colspan=3> " & vbCrLf)
%>
                <input id="Button2" name="Cancel" type="button" class="btn" value="OK" style="WIDTH: 80px" width="80"
                    onclick="self.close()"
                    onmouseover="try{button_onMouseOver(this);}catch(e){}"
                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                    onfocus="try{button_onFocus(this);}catch(e){}"
                    onblur="try{button_onBlur(this);}catch(e){}" />
                <%
                    Response.Write("			    </td>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("			  </tr>" & vbCrLf)
                End If
	
                objExpression = Nothing
	
                Response.Write("<INPUT type=hidden id=txtDisplay name=txtDisplay value=" & fDisplay & ">" & vbCrLf)
                %>
                <tr height="10">
                    <td colspan="5"></td>
                </tr>
            </table>
		</TD>
	</TR>
</table>

</div>
</body>

</html>

<script type="text/javascript">
    util_test_expression_onload();
</script>

<script runat="server" language="vb">

    Function convertLocaleDateToSQL(psDate)
        Dim sLocaleFormat As String
        Dim sSQLFormat As String
        Dim iIndex As Integer
	
        If Len(psDate) > 0 Then
            sLocaleFormat = Session("LocaleDateFormat")
		
            iIndex = InStr(sLocaleFormat, "mm")
            If iIndex > 0 Then
                sSQLFormat = Mid(psDate, iIndex, 2) & "/"
            End If

            iIndex = InStr(sLocaleFormat, "dd")
            If iIndex > 0 Then
                sSQLFormat = sSQLFormat & Mid(psDate, iIndex, 2) & "/"
            End If

            iIndex = InStr(sLocaleFormat, "yyyy")
            If iIndex > 0 Then
                sSQLFormat = sSQLFormat & Mid(psDate, iIndex, 4)
            End If

            convertLocaleDateToSQL = sSQLFormat
        Else
            convertLocaleDateToSQL = ""
        End If
    End Function

</script>
