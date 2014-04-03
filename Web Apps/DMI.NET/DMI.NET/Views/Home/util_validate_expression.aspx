<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<!DOCTYPE html>

<html>
<head>
	<title>OpenHR</title>

	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
	<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-admin-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/fonts/SSI80v194934/style.css")%>" rel="stylesheet" />

	<script type="text/javascript">

		function util_validate_window_onload() {

			if (window.txtDisplay.value != "False") {
				$('#PleaseWaitDiv').hide();
			}
			else
			{
				nextPass();
			}
		}
		
		function overwrite(){
				nextPass();
		}

		function createNew(){
				window.dialogArguments.OpenHR.createNew(self);		
		}

		function removeComponents(piIndex)
		{
				var sKeys;
	
				if (piIndex == 1) {
						sKeys = window.txtDeletedKeys.value;
				}
				else {
						if (piIndex == 2) {
								sKeys = window.txtHiddenNotOwnerKeys.value;
						}
						else {
								sKeys = window.txtHiddenOwnerKeys.value;
						}
				}
				window.dialogArguments.OpenHR.removeComponents(sKeys);		  
				cancelClick();
		}

		function returnToDefSel() {

				window.dialogArguments.OpenHR.returnToDefSel();		  
				cancelClick();
		}

		function makeHidden() {
			document.parentWindow.parent.window.dialogArguments.window.makeHidden(self);
		}

		function nextPass() {
				var iNextPass;

				iNextPass = new Number(frmValidate.validatePass.value);
				iNextPass = iNextPass + 1;

				if (iNextPass <= 3) {
						frmValidate.validatePass.value = iNextPass;
						OpenHR.submitForm(frmValidate);
				}
				else {
						var frmSend = window.dialogArguments.OpenHR.getForm("workframe", "frmSend");
						window.dialogArguments.OpenHR.submitForm(frmSend);
						self.close();
				}
		}

		function cancelClick() {

				var iIndex;
				var sCurrentPage = window.dialogArguments.document.location;

				try {
						sCurrentPage = sCurrentPage.toString();
						iIndex = sCurrentPage.lastIndexOf("/");

						if (iIndex >= 0) {
								sCurrentPage = sCurrentPage.substr(iIndex + 1);
						}
	
						iIndex = sCurrentPage.indexOf(".");
						if (iIndex >= 0) {
								sCurrentPage = sCurrentPage.substr(0, iIndex);
						}

					try {
						sCurrentPage = sCurrentPage.toUpperCase();
					} catch(e) {}
					
					if (sCurrentPage == "UTIL_DEF_EXPRESSION") {
								window.dialogArguments.OpenHR.reEnableControls();
						}
				}
				catch(e) {
				}

				self.close();
		}
	</script>

</head>

<body id="bdyMain">
		
				<div id="util_validate_expression" data-framesource="util_validate_expression" style="text-align: center;">
					<div id="PleaseWaitDiv">
							<h3>
								<%
									If Request.Form("validateUtilType") = 11 Then
										Response.Write("Validating Filter")
									Else
										If Request.Form("validateUtilType") = 12 Then
											Response.Write("Validating Calculation")
										Else
											Response.Write("Validating Expression")
										End If
									End If
%>				
						</h3>
						<br />		
						Loading...
						<br />		
						<br />		
						<input type="button" value="Cancel" class="btn" name="Cancel" style="width: 80px" id="Cancel" onclick="cancelClick()" />
					</div>
<%
	Dim fOK As Boolean
	Dim fDisplay As Boolean
	Dim sUtilType As String
	Dim sUtilType2 As String
	Dim iExprType As Integer

	Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
	Dim objSessionInfo As SessionInfo = CType(Session("SessionContext"), SessionInfo)
	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

	Dim iErrorCode As Integer
	Dim sDeletedKeys As String
	Dim sHiddenOwnerKeys As String
	Dim sHiddenNotOwnerKeys As String
	Dim sDeletedDescs As String
	Dim sHiddenOwnerDescs As String
	Dim sHiddenNotOwnerDescs As String
	Dim iIndex As Integer
	Dim sDesc As String
	
	Dim objExpression As Expression
	Dim iReturnType As Integer
			
	Dim iValidityCode As Integer
	Dim sValidityMessage As String
	Dim iOriginalReturnType As Integer
	Dim sDescription As String
		

	Dim sHiddenErrorMsg As String
		
	fOK = True
	fDisplay = False
	
	If Request.Form("validateUtilType") = "11" Then
		sUtilType = "Filter"
		sUtilType2 = "filter"
		iExprType = 11
	Else
		sUtilType = "Calculation"
		sUtilType2 = "calculation"
		iExprType = 10
	End If
		
	If Request.Form("validatePass") = 1 Then

		Dim prmDeletedKeys = New SqlParameter("psDeletedKeys", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
		Dim prmHiddenOwnerKeys = New SqlParameter("psHiddenOwnerKeys", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
		Dim prmHiddenNotOwnerKeys = New SqlParameter("psHiddenNotOwnerKeys", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
		Dim prmDeletedDescs = New SqlParameter("psDeletedDescs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
		Dim prmHiddenOwnerDescs = New SqlParameter("psHiddenOwnerDescs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
		Dim prmHiddenNotOwnerDescs = New SqlParameter("psHiddenNotOwnerDescs", SqlDbType.VarChar - 1) With {.Direction = ParameterDirection.Output}
		Dim prmErrorCode = New SqlParameter("piErrorCode", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

		Try

			objDataAccess.ExecuteSP("sp_ASRIntValidateExpression" _
				, New SqlParameter("psUtilName", SqlDbType.VarChar, 255) With {.Value = Request.Form("validateName")} _
				, New SqlParameter("piUtilID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("validateUtilID"))} _
				, New SqlParameter("piUtilType", SqlDbType.Int) With {.Value = CleanNumeric(iExprType)} _
				, New SqlParameter("psUtilOwner", SqlDbType.VarChar, 128) With {.Value = Request.Form("validateOwner")} _
				, New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("validateBaseTableID"))} _
				, New SqlParameter("psComponentDefn", SqlDbType.VarChar, -1) With {.Value = Request.Form("components1")} _
				, New SqlParameter("piTimestamp", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("validateTimestamp"))} _
				, prmDeletedKeys, prmHiddenOwnerKeys, prmHiddenNotOwnerKeys, prmDeletedDescs _
				, prmHiddenOwnerDescs, prmHiddenNotOwnerDescs, prmErrorCode)

		Catch ex As Exception
			Throw ex
		End Try

		Response.Write("<input type='hidden' id='txtErrorCode' name='txtErrorCode' value='" & prmErrorCode.Value & "'>" & vbCrLf)
		Response.Write("<input type='hidden' id='txtDeletedKeys' name='txtDeletedKeys' value='" & prmDeletedKeys.Value & "'>" & vbCrLf)
		Response.Write("<input type='hidden' id='txtHiddenOwnerKeys' name='txtHiddenOwnerKeys' value='" & prmHiddenOwnerKeys.Value & "'>" & vbCrLf)
		Response.Write("<input type='hidden' id='txtHiddenNotOwnerKeys' name='txtHiddenNotOwnerKeys' value='" & prmHiddenNotOwnerKeys.Value & "'>" & vbCrLf)

		iErrorCode = CInt(prmErrorCode.Value)
		sDeletedKeys = prmDeletedKeys.Value.ToString()
		sHiddenOwnerKeys = prmHiddenOwnerKeys.Value.ToString()
		sHiddenNotOwnerKeys = prmHiddenNotOwnerKeys.Value.ToString()
		sDeletedDescs = prmDeletedDescs.Value.ToString()
		sHiddenOwnerDescs = prmHiddenOwnerDescs.Value.ToString()
		sHiddenNotOwnerDescs = prmHiddenNotOwnerDescs.Value.ToString()

		If (iErrorCode = 1) Or (iErrorCode = 2) Then
			' 1 = Expression deleted by another user. Save as new ? 
			' 2 = Made hidden/read-only by another user. Save as new ? 
			fDisplay = True

			Response.Write("			  <tr>" & vbCrLf)
			Response.Write("					<td width='20'></td>" & vbCrLf)
			Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
			Response.Write("						<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
			Response.Write("			    </td>" & vbCrLf)
			Response.Write("					<td width='20'></td>" & vbCrLf)
			Response.Write("			  </tr>" & vbCrLf)
			Response.Write("			  <tr>" & vbCrLf)
			Response.Write("					<td width='20'></td>" & vbCrLf)
			Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)

			If (iErrorCode = 1) Then
				Response.Write("						The " & sUtilType2 & " has been deleted by another user. Save as a new definition ?" & vbCrLf)
			Else
				Response.Write("						The " & sUtilType2 & " has been amended by another user and is now Read Only. Save as a new definition ?" & vbCrLf)
			End If

			Response.Write("			    </td>" & vbCrLf)
			Response.Write("					<td width='20'></td>" & vbCrLf)
			Response.Write("			  </tr>" & vbCrLf)
			Response.Write("			  <tr>" & vbCrLf)
			Response.Write("					<td height='20' colspan='5'></td>" & vbCrLf)
			Response.Write("			  </tr>" & vbCrLf)
			Response.Write("			  <tr> " & vbCrLf)
			Response.Write("					<td width='20'></td>" & vbCrLf)
			Response.Write("			    <td align='right'> " & vbCrLf)
			Response.Write("    				    <input type='button' value='Yes' class='btn' name='btnYes' style='width: 80px; ' id='btnYes'" & vbCrLf)
			Response.Write("    				        OnClick=""createNew()""/>" & vbCrLf)
			Response.Write("			    </td>" & vbCrLf)
			Response.Write("					<td width='20'></td>" & vbCrLf)
			Response.Write("			    <td align='left'> " & vbCrLf)
			Response.Write("    				    <input type='button' value='No' class='btn' name='btnNo' style='width: 80px' id='btnNo'" & vbCrLf)
			Response.Write("    				        OnClick=""cancelClick()""/>" & vbCrLf)
			Response.Write("			    </td>" & vbCrLf)
			Response.Write("					<td width='20'></td>" & vbCrLf)
			Response.Write("			  </tr>" & vbCrLf)
		End If
	
		If iErrorCode = 4 Then
			' 4 = Non-unique name. Save fails */
			fDisplay = True
			
			Response.Write("	<div style='text-align: center'>")
			Response.Write("		<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
			Response.Write("		A " & sUtilType2 & " called '" & Request.Form("validateName") & "' already exists.<br/><br/>")
			Response.Write("    <input type='button' value='Close' class='btn' name='Cancel' style='width: 80px' id='Cancel' OnClick='cancelClick()'/>" & vbCrLf)
			Response.Write("	</div>")
		End If

		If (iErrorCode = 0) Or (iErrorCode = 3) Then
			' 0 = No error (but must check the strings of keys)
			If Len(sDeletedKeys) > 0 Then
				fDisplay = True
				
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
				Response.Write("						<H3>Error Saving " & sUtilType & "</H3>" & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
				Response.Write("						 The following calculations and filters have been deleted, and will be removed from the " & sUtilType2 & " definition." & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("					<td height='20' colspan='5'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				
				iIndex = InStr(sDeletedDescs, "	")
				Do While iIndex > 0
					sDesc = Left(sDeletedDescs, iIndex - 1)

					Response.Write("			  <tr>" & vbCrLf)
					Response.Write("					<td width='20'></td>" & vbCrLf)
					Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
					Response.Write("						 " & sDesc & vbCrLf)
					Response.Write("			    </td>" & vbCrLf)
					Response.Write("					<td width='20'></td>" & vbCrLf)
					Response.Write("			  </tr>" & vbCrLf)
					
					sDeletedDescs = Mid(sDeletedDescs, iIndex + 1)
					iIndex = InStr(sDeletedDescs, "	")
				Loop

				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
				Response.Write("						 " & sDeletedDescs & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)

				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("					<td height='20' colspan='5'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr> " & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
				Response.Write("    				    <input type='button' value='Close' class='btn' name='RemoveComponents' style='width: 80px' id='RemoveComponents'" & vbCrLf)
				Response.Write("    				        OnClick=""removeComponents(1)""/>" & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
			Else
				If (Len(sHiddenOwnerKeys) > 0) Or (Len(sHiddenNotOwnerKeys) > 0) Then

					If (UCase(Request.Form("validateOwner")) = UCase(Session("Username"))) Then
						' Current user IS the owner of the filter/calc.
						If (Len(sHiddenNotOwnerKeys) > 0) Then
							' There are hidden components in the expression that are NOT owned by the current user.
							' Need to remove the hidden components.	
							fDisplay = True
							
							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
							Response.Write("						<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
							Response.Write("			    </td>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
							Response.Write("						 The following calculations and filters have been made hidden, and will be removed from the " & sUtilType2 & " definition." & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td height='20' colspan='5'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
										
							iIndex = InStr(sHiddenNotOwnerDescs, "	")
							Do While iIndex > 0
								sDesc = Left(sHiddenNotOwnerDescs, iIndex - 1)

								Response.Write("			  <tr>" & vbCrLf)
								Response.Write("					<td width='20'></td>" & vbCrLf)
								Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
								Response.Write("						 " & sDesc & vbCrLf)
								Response.Write("			    </td>" & vbCrLf)
								Response.Write("					<td width='20'></td>" & vbCrLf)
								Response.Write("			  </tr>" & vbCrLf)
											
								sHiddenNotOwnerDescs = Mid(sHiddenNotOwnerDescs, iIndex + 1)
								iIndex = InStr(sHiddenNotOwnerDescs, "	")
							Loop

							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
							Response.Write("						 " & sHiddenNotOwnerDescs & vbCrLf)
							Response.Write("			    </td>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)

							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td height='20' colspan='5'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
							Response.Write("			  <tr> " & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
							Response.Write("    				    <input type='button' value='Close' class='btn' name='RemoveComponents' style='width: 80px' id='RemoveComponents'" & vbCrLf)
							Response.Write("    				        OnClick=""removeComponents(2)""/>" & vbCrLf)
							Response.Write("			    </td>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
						Else
							' There are hidden components in the expression that ARE owned by the current user.
							' Need to make the expression hidden too.
							If (Request.Form("validateAccess") <> "HD") Then
								fDisplay = True
								
								Response.Write("<div style='text-align: center'>")
								Response.Write("<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
								Response.Write("The following calculations and filters have been made hidden. The " & sUtilType2 & " will now be made hidden.<br/><br/>")
								
								iIndex = InStr(sHiddenOwnerDescs, "	")
								Do While iIndex > 0
									sDesc = Left(sHiddenOwnerDescs, iIndex - 1)
									Response.Write(sDesc & "<br/>")
									sHiddenOwnerDescs = Mid(sHiddenOwnerDescs, iIndex + 1)
									iIndex = InStr(sHiddenOwnerDescs, "	")
								Loop

								Response.Write(sHiddenOwnerDescs & "<br/><br/>")
								Response.Write("<input type='button' value='Close' class='btn' name='makeHidden' style='width: 80px' id='makeHidden' OnClick='makeHidden()'/>" & vbCrLf)
								Response.Write("</div>")
								
							End If
						End If
					Else
						' Current user is NOT the owner of the filter/calc.
						fDisplay = True

						If (Len(sHiddenNotOwnerKeys) > 0) Then
							' There are hidden components in the expression that are NOT owned by the current user.
							' Cannot edit the expression as it too must now be hidden.
							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
							Response.Write("						<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
							Response.Write("			    </td>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
							Response.Write("						 The following calculations and filters have been made hidden by another user. Cannot make any modifications to this " & sUtilType2 & " definition." & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td height='20' colspan='5'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
										
							iIndex = InStr(sHiddenNotOwnerDescs, "	")
							Do While iIndex > 0
								sDesc = Left(sHiddenNotOwnerDescs, iIndex - 1)

								Response.Write("			  <tr>" & vbCrLf)
								Response.Write("					<td width='20'></td>" & vbCrLf)
								Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
								Response.Write("						 " & sDesc & vbCrLf)
								Response.Write("			    </td>" & vbCrLf)
								Response.Write("					<td width='20'></td>" & vbCrLf)
								Response.Write("			  </tr>" & vbCrLf)
											
								sHiddenNotOwnerDescs = Mid(sHiddenNotOwnerDescs, iIndex + 1)
								iIndex = InStr(sHiddenNotOwnerDescs, "	")
							Loop

							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
							Response.Write("						 " & sHiddenNotOwnerDescs & vbCrLf)
							Response.Write("			    </td>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)

							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td height='20' colspan='5'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
							Response.Write("			  <tr> " & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
							Response.Write("    				    <input type='button' value='Close' class='btn' name='ReturnToDefSel' style='width: 80px' id='ReturnToDefSel'" & vbCrLf)
							Response.Write("    				        OnClick=""returnToDefSel()""/>" & vbCrLf)
							Response.Write("			    </td>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
						Else
							' There are hidden components in the expression that ARE owned by the current user.
							' Need to remove the hidden components.	
							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
							Response.Write("						<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
							Response.Write("			    </td>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
							Response.Write("						 The following calculations and filters have been made hidden, and will be removed from the " & sUtilType2 & " definition." & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td height='20' colspan='5'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
										
							iIndex = InStr(sHiddenOwnerDescs, "	")
							Do While iIndex > 0
								sDesc = Left(sHiddenOwnerDescs, iIndex - 1)

								Response.Write("			  <tr>" & vbCrLf)
								Response.Write("					<td width='20'></td>" & vbCrLf)
								Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
								Response.Write("						 " & sDesc & vbCrLf)
								Response.Write("			    </td>" & vbCrLf)
								Response.Write("					<td width='20'></td>" & vbCrLf)
								Response.Write("			  </tr>" & vbCrLf)
											
								sHiddenOwnerDescs = Mid(sHiddenOwnerDescs, iIndex + 1)
								iIndex = InStr(sHiddenOwnerDescs, "	")
							Loop

							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
							Response.Write("						 " & sHiddenOwnerDescs & vbCrLf)
							Response.Write("			    </td>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)

							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td height='20' colspan='5'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
							Response.Write("			  <tr> " & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
							Response.Write("    				    <input type='button' value='Close' class='btn' name='RemoveComponents' style='width: 80px' id='RemoveComponents'" & vbCrLf)
							Response.Write("    				        OnClick=""removeComponents(3)""/>" & vbCrLf)
							Response.Write("			    </td>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
						End If
					End If
				End If
			End If
		End If
	
		If (iErrorCode = 3) And (fDisplay = False) Then
			' 3 = Modified by another user (still writable). Overwrite ? 
			fDisplay = True

			Response.Write("			  <tr>" & vbCrLf)
			Response.Write("					<td width='20'></td>" & vbCrLf)
			Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
			Response.Write("						<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
			Response.Write("			    </td>" & vbCrLf)
			Response.Write("					<td width='20'></td>" & vbCrLf)
			Response.Write("			  </tr>" & vbCrLf)
			Response.Write("			  <tr>" & vbCrLf)
			Response.Write("					<td width='20'></td>" & vbCrLf)
			Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
			Response.Write("						The " & sUtilType2 & " has been amended by another user. Would you like to overwrite this definition ?" & vbCrLf)
			Response.Write("			    </td>" & vbCrLf)
			Response.Write("					<td width='20'></td>" & vbCrLf)
			Response.Write("			  </tr>" & vbCrLf)
			Response.Write("			  <tr>" & vbCrLf)
			Response.Write("					<td height='20' colspan='5'></td>" & vbCrLf)
			Response.Write("			  </tr>" & vbCrLf)
			Response.Write("			  <tr> " & vbCrLf)
			Response.Write("					<td width='20'></td>" & vbCrLf)
			Response.Write("			    <td align='right'> " & vbCrLf)
			Response.Write("    				    <input type='button' value='Yes' class='btn' name='btnYes' style='width: 80px' id='btnYes'" & vbCrLf)
			Response.Write("    				        OnClick=""overwrite()""/>" & vbCrLf)
			Response.Write("			    </td>" & vbCrLf)
			Response.Write("					<td width='20'></td>" & vbCrLf)
			Response.Write("			    <td align='left'> " & vbCrLf)
			Response.Write("    				    <input type='button' value='No' class='btn' name='btnNo' style='width: 80px' id='btnNo'" & vbCrLf)
			Response.Write("    				        OnClick=""cancelClick()""/>" & vbCrLf)
			Response.Write("			    </td>" & vbCrLf)
			Response.Write("					<td width='20'></td>" & vbCrLf)
			Response.Write("				</tr>" & vbCrLf)
		End If
	End If
	
		
	If Request.Form("validatePass") = 2 Then
		' Get the server DLL to validate the expression definition
		objExpression = New Expression(objSessionInfo)
			
		If Request.Form("validateUtilType") = 11 Then
			iExprType = 11
			iReturnType = 3
		Else
			iExprType = 10
			iReturnType = 0
		End If
				
		fOK = objExpression.Initialise(CLng(Request.Form("validateBaseTableID")), CLng(Request.Form("validateUtilID")), CInt(iExprType), CInt(iReturnType))

		If fOK Then
			fOK = objExpression.SetExpressionDefinition(CStr(Request.Form("components1")), "", "", "", "", "")
		End If

		If fOK Then
			iValidityCode = objExpression.ValidateExpression
			If iValidityCode > 0 Then
				fDisplay = True
				Response.Write("			  <table align = 'center'>" & vbCrLf)
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			        <td align='center' colspan='3'> " & vbCrLf)
				Response.Write("					    <h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
				Response.Write("			        </td>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			        <td align='center' colspan='3'> " & vbCrLf)
				sValidityMessage = objExpression.ValidityMessage(CInt(iValidityCode))
				sValidityMessage = Replace(sValidityMessage, vbCr, "<BR>")
				Response.Write("						 " & sValidityMessage & vbCrLf)
				Response.Write("			        </td>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("					<td height='20'></td>" & vbCrLf)
				Response.Write("					<td height='20'  colspan='3'></td>" & vbCrLf)
				Response.Write("					<td height='20' ></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr> " & vbCrLf)
				Response.Write("				<td width='20'></td>" & vbCrLf)
				Response.Write("			    <td align='center'  colspan='3'> " & vbCrLf)
				Response.Write("    				    <input type='button' value='Close' class='btn' name='Cancel' style='width: 80px' id='Cancel'" & vbCrLf)
				Response.Write("    				        OnClick=""cancelClick()""/>" & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  </table>" & vbCrLf)
			Else
				iReturnType = objExpression.ReturnType
			End If
		End If
			
		objExpression = Nothing
	
		If Not fDisplay Then
			' Check if the expression return type has changed. 
			' If so, check if it can be.

			If Request.Form("validateUtilID") > 0 Then
				objExpression = New Expression(objSessionInfo)

				iOriginalReturnType = objExpression.ExistingExpressionReturnType(CLng(Request.Form("validateUtilID")))
				objExpression = Nothing
					
				If iReturnType <> iOriginalReturnType Then
										
					Dim rsUsage = objDatabase.GetUtilityUsage(CInt(CleanNumeric(Request.Form("validateUtilType"))), CInt(CleanNumeric(Request.Form("validateUtilID"))))

					If rsUsage.Rows.Count = 0 Then
						Response.Write("<option>&lt;None&gt;</option>")
					Else
						For Each objRow As DataRow In rsUsage.Rows
							sDescription = objRow("description").ToString()
							sDescription = Replace(sDescription, "<", "&lt;")
							sDescription = Replace(sDescription, ">", "&gt;")
							Response.Write("<option>" & sDescription & "</option>")
						Next
					End If
					 
					If rsUsage.Rows.Count > 0 Then
						fDisplay = True
						Response.Write("			  <tr>" & vbCrLf)
						Response.Write("					<td width='20'></td>" & vbCrLf)
						Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
						Response.Write("						<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
						Response.Write("			    </td>" & vbCrLf)
						Response.Write("					<td width='20'></td>" & vbCrLf)
						Response.Write("			  </tr>" & vbCrLf)
						Response.Write("			  <tr>" & vbCrLf)
						Response.Write("					<td width='20'></td>" & vbCrLf)
						Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
						Response.Write("						 The return type cannot be changed, as the " & sUtilType2 & " is currently being used in the following definitions." & vbCrLf)
						Response.Write("			    </td>" & vbCrLf)
						Response.Write("					<td width='20'></td>" & vbCrLf)
						Response.Write("			  </tr>" & vbCrLf)
						Response.Write("			  <tr>" & vbCrLf)
						Response.Write("					<td height='20' colspan='5'></td>" & vbCrLf)
						Response.Write("			  </tr>" & vbCrLf)
						
						For Each objRow As DataRow In rsUsage.Rows
							sDescription = objRow("description").ToString()
							sDescription = Replace(sDescription, "<", "&lt;")
							sDescription = Replace(sDescription, ">", "&gt;")

							Response.Write("			  <tr>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
							Response.Write("						 " & sDescription & vbCrLf)
							Response.Write("			    </td>" & vbCrLf)
							Response.Write("					<td width='20'></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
						Next
						
						Response.Write("			  <tr>" & vbCrLf)
						Response.Write("					<td height='20' colspan='5'></td>" & vbCrLf)
						Response.Write("			  </tr>" & vbCrLf)
						Response.Write("			  <tr> " & vbCrLf)
						Response.Write("					<td width='20'></td>" & vbCrLf)
						Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
						Response.Write("    				    <input type='button' value='Close' class='btn' name='Cancel' style='width: 80px' id='Cancel'" & vbCrLf)
						Response.Write("    				        OnClick=""cancelClick()""/>" & vbCrLf)
						Response.Write("			    </td>" & vbCrLf)
						Response.Write("					<td width='20'></td>" & vbCrLf)
						Response.Write("			  </tr>" & vbCrLf)
					End If
										 
				End If
			End If
		End If
	End If
		
	If Request.Form("validatePass") = 3 Then
		If (Request.Form("validateUtilID") > 0) And _
				(UCase(Request.Form("validateOwner")) = UCase(Session("username"))) And _
				(Request.Form("validateAccess") = "HD") And _
				(Request.Form("validateOriginalAccess") <> "HD") Then
			' Check if the expression can be made hidden.

			Dim prmResult = New SqlParameter("piResult", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmMsg = New SqlParameter("psMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

			objDataAccess.ExecuteSP("sp_ASRIntCheckCanMakeHidden" _
							, New SqlParameter("piUtilityType", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("validateUtilType"))} _
							, New SqlParameter("piUtilityID", SqlDbType.VarChar, 255) With {.Value = CleanNumeric(Request.Form("validateUtilID"))} _
							, prmResult, prmMsg)

			If prmResult.Value = 1 Then
				' calc/filter used only in utilities owned by the current user - we then need to prompt the user if they want to make these utilities hidden too.
				fDisplay = True
				sHiddenErrorMsg = "Making this " & sUtilType2 & " hidden will automatically make the following definition(s), of which you are the owner, hidden also :" & _
					"<BR><BR>" & _
					prmMsg.Value.ToString() & _
					"<BR><BR>" & _
					"Do you wish to continue ?"
				
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
				Response.Write("						<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
				Response.Write("						 " & sHiddenErrorMsg & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("					<td height='20' colspan='5'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr> " & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			    <td align='right'> " & vbCrLf)
				Response.Write("    				    <input type='button' value='Yes' class='btn' name='btnYes' style='width: 80px' id='btnYes'" & vbCrLf)
				Response.Write("    				        OnClick=""overwrite()""/>" & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			    <td align='left'> " & vbCrLf)
				Response.Write("    				    <input type='button' value='No' class='btn' name='btnNo' style='width: 80px' id='btnNo'" & vbCrLf)
				Response.Write("    				        OnClick=""cancelClick()""/>" & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("				</tr>" & vbCrLf)
			End If

			If (prmResult.Value = 2) Or _
				(prmResult.Value = 3) Or _
				(prmResult.Value = 4) Then
				' calc/filter used in utilities which are in batch jobs not owned by the current user - Cannot therefore make the calc/filter hidden.
				If (prmResult.Value = 2) Then
					sHiddenErrorMsg = "This " & sUtilType2 & " cannot be made hidden as it is used in definition(s) which are included in the following batch jobs of which you are not the owner :" & _
						"<BR><BR>" & _
						prmMsg.Value.ToString()
				Else
					If (prmResult.Value = 3) Then
						sHiddenErrorMsg = "This " & sUtilType2 & " cannot be made hidden as it is used in definition(s), of which you are not the owner :" & _
							"<BR><BR>" & _
							prmMsg.Value.ToString()
					Else
						sHiddenErrorMsg = "This " & sUtilType2 & " cannot be made hidden as it is used in definition(s) which are included in the following batch jobs which are scheduled to be run by other user groups :" & _
							"<BR><BR>" & _
							prmMsg.Value.ToString()
					End If
				End If
				fDisplay = True
				
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
				Response.Write("						<h3>Error Saving " & sUtilType & "</h3>" & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
				Response.Write("						 " & sHiddenErrorMsg & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("					<td height='20' colspan='5></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr> " & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			    <td align='center' colspan='3'> " & vbCrLf)
				Response.Write("    				    <input type='button' value='Close' class='btn' name='Cancel' style='width: 80px' id='Cancel'" & vbCrLf)
				Response.Write("    				        OnClick=""cancelClick()""/>" & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("					<td width='20'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
			End If
		End If
	End If
	
	Response.Write("<input type='hidden' id='txtDisplay' name='txtDisplay' value='" & fDisplay & "'>" & vbCrLf)
%>
					<tr height="10">
						<td colspan="5"></td>
					</tr>
					</table>
		</td>
	</tr>
</table>
</div>
						

		<form id="frmValidate" name="frmValidate" method="post" action="util_validate_expression" style="visibility: hidden; display: none">
				<input type="hidden" id="validatePass" name="validatePass" value='<%=Request.form("validatePass")%>'>
				<input type="hidden" id="validateUtilID" name="validateUtilID" value='<%=Request.form("validateUtilID")%>'>
				<input type="hidden" id="validateUtilType" name="validateUtilType" value='<%=Request.form("validateUtilType")%>'>
				<input type="hidden" id="validateAccess" name="validateAccess" value='<%=Request.form("validateAccess")%>'>
				<input type="hidden" id="validateOriginalAccess" name="validateOriginalAccess" value='<%=Request.form("validateOriginalAccess")%>'>
				<input type="hidden" id="validateOwner" name="validateOwner" value='<%=Request.form("validateOwner")%>'>

				<input type="hidden" id="components1" name="components1" value="<%=Request.form("components1")%>">
				<input type="hidden" id="validateBaseTableID" name="validateBaseTableID" value='<%=Request.form("validateBaseTableID")%>'>
		</form>


</body>

</html>

<script type="text/javascript">
		util_validate_window_onload();
</script>
