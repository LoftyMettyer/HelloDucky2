<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>

<html>
<head>
    <title>OpenHR Intranet</title>
    

<script type="text/javascript">
<!--
    
    function util_validate_window_onload() {

        debugger;

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
            sKeys = txtDeletedKeys.value;
        }
        else {
            if (piIndex == 2) {
                sKeys = txtHiddenNotOwnerKeys.value;
            }
            else {
                sKeys = txtHiddenOwnerKeys.value;
            }
        }
        window.dialogArguments.OpenHR.removeComponents(sKeys);		  
        cancelClick();
    }

    function returnToDefSel() {

        window.dialogArguments.OpenHR.returnToDefSel();		  
        cancelClick();
    }

    function makeHidden() 
    {
        window.dialogArguments.OpenHR.makeHidden(self);
    }

    function nextPass()
    {
        var iNextPass;
        var sURL;
	
        iNextPass = new Number(frmValidate.validatePass.value);
        iNextPass = iNextPass + 1;

        if (iNextPass <= 3) {
            frmValidate.validatePass.value = iNextPass;

            OpenHR.submitForm(frmValidate);
        }
        else {	
            
            OpenHR.submitForm(window.dialogArguments.document.getElementById('frmSend'));
            self.close();
        }
    }

    function cancelClick()
    {
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
	
            sCurrentPage = sCurrentPage.toUpperCase();

            if (sCurrentPage == "UTIL_DEF_EXPRESSION") {
                window.dialogArguments.OpenHR.reEnableControls();		  
            }
        }
        catch(e) {
        }

        self.close();
    }
    -->
</script>

    

</head>
<body id=bdyMain >
    
        <div id="util_validate_expression" data-framesource="util_validate_expression">

    
<table align=center class="outline" cellPadding=5 cellSpacing=0/>
	<TR>
		<TD>
			<table class="invisible" cellspacing="0" cellpadding="0"/>
				<tr> 
			    <td colspan=5 height=10></td>
			  </tr>

			  <tr id=trPleaseWait1> 
					<td width=20></td>
			    <td align=center colspan=3> 
<%
	if Request.form("validateUtilType") = 11 then
        Response.Write("Validating Filter")
    Else
        If Request.Form("validateUtilType") = 12 Then
            Response.Write("Validating Calculation")
        Else
            Response.Write("Validating Expression")
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
						<INPUT TYPE=button VALUE="Cancel" class="btn" NAME="Cancel" style="WIDTH: 80px" width=80 id=Cancel
						    OnClick="cancelClick()" 
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
	dim sUtilType2
	dim iExprType

    Dim cmdValidate
    Dim prmUtilName
    Dim prmUtilID
    Dim prmExprType
    Dim prmUtilOwner
    Dim prmBaseTableID
    Dim prmComponentDefn
    Dim prmTimestamp
    Dim prmDeletedKeys
    Dim prmHiddenOwnerKeys
    Dim prmHiddenNotOwnerKeys
    Dim prmDeletedDescs
    Dim prmHiddenOwnerDescs
    Dim prmHiddenNotOwnerDescs
    Dim prmErrorCode

    Dim iErrorCode As Integer
    Dim sDeletedKeys As String
    Dim sHiddenOwnerKeys As String
    Dim sHiddenNotOwnerKeys As String
    Dim sDeletedDescs As String
    Dim sHiddenOwnerDescs As String
    Dim sHiddenNotOwnerDescs As String
    Dim iIndex As Integer
    Dim sDesc As String
  
    Dim objExpression
    Dim iReturnType As Integer
      
    Dim iValidityCode As Integer
    Dim sValidityMessage As String
    Dim iOriginalReturnType As Integer
    Dim cmdDefPropRecords
    Dim prmType
    Dim prmID
    Dim rsDefProp
    Dim sDescription As String
    Dim cmdCheckHidden
    
    Dim prmResult
    Dim prmMsg
    Dim sHiddenErrorMsg As String
    
	fOK = true
	fDisplay = false
	
	if Request.form("validateUtilType") = "11" then
		sUtilType = "Filter"
		sUtilType2 = "filter"
		iExprType = 11
	else
		sUtilType = "Calculation"
		sUtilType2 = "calculation"
		iExprType = 10
	end if
		
	if Request.form("validatePass") = 1 then
        cmdValidate = CreateObject("ADODB.Command")
		cmdValidate.CommandText = "sp_ASRIntValidateExpression"
		cmdValidate.CommandType = 4 ' Stored Procedure
        cmdValidate.ActiveConnection = Session("databaseConnection")

        prmUtilName = cmdValidate.CreateParameter("utilName", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdValidate.Parameters.Append(prmUtilName)
		prmUtilName.value = Request.form("validateName")

        prmUtilID = cmdValidate.CreateParameter("utilID", 3, 1) '3=integer, 1=input
        cmdValidate.Parameters.Append(prmUtilID)
		prmUtilID.value = cleanNumeric(Request.form("validateUtilID"))

        prmExprType = cmdValidate.CreateParameter("exprtype", 3, 1) '3=integer, 1=input
        cmdValidate.Parameters.Append(prmExprType)
		prmExprType.value = cleanNumeric(iExprType)

        prmUtilOwner = cmdValidate.CreateParameter("utilOwner", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdValidate.Parameters.Append(prmUtilOwner)
		prmUtilOwner.value = Request.form("validateOwner")

        prmBaseTableID = cmdValidate.CreateParameter("baseTableID", 3, 1) '3=integer, 1=input
        cmdValidate.Parameters.Append(prmBaseTableID)
		prmBaseTableID.value = cleanNumeric(Request.form("validateBaseTableID"))

        prmComponentDefn = cmdValidate.CreateParameter("componentDefn", 200, 1, 2147483646)
        cmdValidate.Parameters.Append(prmComponentDefn)
		prmComponentDefn.value = Request.form("components1")

        prmTimestamp = cmdValidate.CreateParameter("timestamp", 3, 1) '3=integer, 1=input
        cmdValidate.Parameters.Append(prmTimestamp)
		prmTimestamp.value = cleanNumeric(Request.form("validateTimestamp"))

        prmDeletedKeys = cmdValidate.CreateParameter("deletedKeys", 200, 2, 2147483646)
        cmdValidate.Parameters.Append(prmDeletedKeys)

        prmHiddenOwnerKeys = cmdValidate.CreateParameter("hiddenOwnerKeys", 200, 2, 2147483646)
        cmdValidate.Parameters.Append(prmHiddenOwnerKeys)

        prmHiddenNotOwnerKeys = cmdValidate.CreateParameter("hiddenNotOwnerKeys", 200, 2, 2147483646)
        cmdValidate.Parameters.Append(prmHiddenNotOwnerKeys)
	
        prmDeletedDescs = cmdValidate.CreateParameter("deletedDescs", 200, 2, 2147483646)
        cmdValidate.Parameters.Append(prmDeletedDescs)

        prmHiddenOwnerDescs = cmdValidate.CreateParameter("hiddenOwnerDescs", 200, 2, 2147483646)
        cmdValidate.Parameters.Append(prmHiddenOwnerDescs)

        prmHiddenNotOwnerDescs = cmdValidate.CreateParameter("hiddenNotOwnerDescs", 200, 2, 2147483646)
        cmdValidate.Parameters.Append(prmHiddenNotOwnerDescs)

        prmErrorCode = cmdValidate.CreateParameter("errorCode", 3, 2) '3=integer, 2=output
        cmdValidate.Parameters.Append(prmErrorCode)

        Err.Clear()
		cmdValidate.Execute

        Response.Write("<INPUT type=hidden id=txtErrorCode name=txtErrorCode value=" & cmdValidate.Parameters("errorCode").Value & ">" & vbCrLf)
        Response.Write("<INPUT type=hidden id=txtDeletedKeys name=txtDeletedKeys value=""" & cmdValidate.Parameters("deletedKeys").Value & """>" & vbCrLf)
        Response.Write("<INPUT type=hidden id=txtHiddenOwnerKeys name=txtHiddenOwnerKeys value=""" & cmdValidate.Parameters("hiddenOwnerKeys").Value & """>" & vbCrLf)
        Response.Write("<INPUT type=hidden id=txtHiddenNotOwnerKeys name=txtHiddenNotOwnerKeys value=""" & cmdValidate.Parameters("hiddenNotOwnerKeys").Value & """>" & vbCrLf)

        iErrorCode = cmdValidate.Parameters("errorCode").Value
		sDeletedKeys = cmdValidate.Parameters("deletedKeys").Value
		sHiddenOwnerKeys = cmdValidate.Parameters("hiddenOwnerKeys").Value
		sHiddenNotOwnerKeys = cmdValidate.Parameters("hiddenNotOwnerKeys").Value
		sDeletedDescs = cmdValidate.Parameters("deletedDescs").Value
		sHiddenOwnerDescs = cmdValidate.Parameters("hiddenOwnerDescs").Value
		sHiddenNotOwnerDescs = cmdValidate.Parameters("hiddenNotOwnerDescs").Value

        cmdValidate = Nothing
											
		if (iErrorCode = 1) or _
			(iErrorCode = 2) then
			' 1 = Expression deleted by another user. Save as new ? 
			' 2 = Made hidden/read-only by another user. Save as new ? 
			fDisplay = true

            Response.Write("			  <tr>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
            Response.Write("						<H3>Error Saving " & sUtilType & "</H3>" & vbCrLf)
            Response.Write("			    </td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			  </tr>" & vbCrLf)
            Response.Write("			  <tr>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=center colspan=3> " & vbCrLf)

            If (iErrorCode = 1) Then
                Response.Write("						The " & sUtilType2 & " has been deleted by another user. Save as a new definition ?" & vbCrLf)
            Else
                Response.Write("						The " & sUtilType2 & " has been amended by another user and is now Read Only. Save as a new definition ?" & vbCrLf)
            End If

            Response.Write("			    </td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			  </tr>" & vbCrLf)
            Response.Write("			  <tr>" & vbCrLf)
            Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
            Response.Write("			  </tr>" & vbCrLf)
            Response.Write("			  <tr> " & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=right> " & vbCrLf)
            Response.Write("    				    <INPUT TYPE=button VALUE=Yes class=""btn"" NAME=btnYes style=""WIDTH: 80px"" width=80 id=btnYes" & vbCrLf)
            Response.Write("    				        OnClick=""createNew()""" & vbCrLf)
            Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
            Response.Write("			    </td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=left> " & vbCrLf)
            Response.Write("    				    <INPUT TYPE=button VALUE=No class=""btn"" NAME=btnYes style=""WIDTH: 80px"" width=80 id=btnNo" & vbCrLf)
            Response.Write("    				        OnClick=""cancelClick()""" & vbCrLf)
            Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
            Response.Write("			    </td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			  </tr>" & vbCrLf)
        End If
	
        If iErrorCode = 4 Then
            ' 4 = Non-unique name. Save fails */
            fDisplay = True
			
            Response.Write("			  <tr>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
            Response.Write("						<H3>Error Saving " & sUtilType & "</H3>" & vbCrLf)
            Response.Write("			    </td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			  </tr>" & vbCrLf)
            Response.Write("			  <tr>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
            Response.Write("						 A " & sUtilType2 & " called '" & Request.Form("validateName") & "' already exists." & vbCrLf)
            Response.Write("			    </td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			  </tr>" & vbCrLf)
            Response.Write("			  <tr>" & vbCrLf)
            Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
            Response.Write("			  </tr>" & vbCrLf)
            Response.Write("			  <tr> " & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
            Response.Write("    				    <INPUT TYPE=button VALUE=Close class=""btn"" NAME=Cancel style=""WIDTH: 80px"" width=80 id=Cancel" & vbCrLf)
            Response.Write("    				        OnClick=""cancelClick()""" & vbCrLf)
            Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
            Response.Write("			    </td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			  </tr>" & vbCrLf)
        End If

        If (iErrorCode = 0) Or _
            (iErrorCode = 3) Then
			' 0 = No error (but must check the strings of keys)
			if len(sDeletedKeys) > 0 then
				fDisplay = true
				
                Response.Write("			  <tr>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                Response.Write("						<H3>Error Saving " & sUtilType & "</H3>" & vbCrLf)
                Response.Write("			    </td>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			  </tr>" & vbCrLf)
                Response.Write("			  <tr>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                Response.Write("						 The following calculations and filters have been deleted, and will be removed from the " & sUtilType2 & " definition." & vbCrLf)
                Response.Write("			    </td>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			  </tr>" & vbCrLf)
                Response.Write("			  <tr>" & vbCrLf)
                Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                Response.Write("			  </tr>" & vbCrLf)
				
                iIndex = InStr(sDeletedDescs, "	")
                Do While iIndex > 0
                    sDesc = Left(sDeletedDescs, iIndex - 1)

                    Response.Write("			  <tr>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                    Response.Write("						 " & sDesc & vbCrLf)
                    Response.Write("			    </td>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("			  </tr>" & vbCrLf)
					
                    sDeletedDescs = Mid(sDeletedDescs, iIndex + 1)
                    iIndex = InStr(sDeletedDescs, "	")
				loop

                Response.Write("			  <tr>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                Response.Write("						 " & sDeletedDescs & vbCrLf)
                Response.Write("			    </td>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			  </tr>" & vbCrLf)

                Response.Write("			  <tr>" & vbCrLf)
                Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                Response.Write("			  </tr>" & vbCrLf)
                Response.Write("			  <tr> " & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                Response.Write("    				    <INPUT TYPE=button VALUE=Close class=""btn"" NAME=RemoveComponents style=""WIDTH: 80px"" width=80 id=RemoveComponents" & vbCrLf)
                Response.Write("    				        OnClick=""removeComponents(1)""" & vbCrLf)
                Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
                Response.Write("			    </td>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			  </tr>" & vbCrLf)
            Else
                If (Len(sHiddenOwnerKeys) > 0) Or _
                    (Len(sHiddenNotOwnerKeys) > 0) Then

                    If (UCase(Request.Form("validateOwner")) = UCase(Session("Username"))) Then
                        ' Current user IS the owner of the filter/calc.
                        If (Len(sHiddenNotOwnerKeys) > 0) Then
                            ' There are hidden components in the expression that are NOT owned by the current user.
                            ' Need to remove the hidden components.	
                            fDisplay = True
							
                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                            Response.Write("						<H3>Error Saving " & sUtilType & "</H3>" & vbCrLf)
                            Response.Write("			    </td>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                            Response.Write("						 The following calculations and filters have been made hidden, and will be removed from the " & sUtilType2 & " definition." & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
										
                            iIndex = InStr(sHiddenNotOwnerDescs, "	")
                            Do While iIndex > 0
                                sDesc = Left(sHiddenNotOwnerDescs, iIndex - 1)

                                Response.Write("			  <tr>" & vbCrLf)
                                Response.Write("					<td width=20></td>" & vbCrLf)
                                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                                Response.Write("						 " & sDesc & vbCrLf)
                                Response.Write("			    </td>" & vbCrLf)
                                Response.Write("					<td width=20></td>" & vbCrLf)
                                Response.Write("			  </tr>" & vbCrLf)
											
                                sHiddenNotOwnerDescs = Mid(sHiddenNotOwnerDescs, iIndex + 1)
                                iIndex = InStr(sHiddenNotOwnerDescs, "	")
                            Loop

                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                            Response.Write("						 " & sHiddenNotOwnerDescs & vbCrLf)
                            Response.Write("			    </td>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)

                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
                            Response.Write("			  <tr> " & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                            Response.Write("    				    <INPUT TYPE=button VALUE=Close class=""btn"" NAME=RemoveComponents style=""WIDTH: 80px"" width=80 id=RemoveComponents" & vbCrLf)
                            Response.Write("    				        OnClick=""removeComponents(2)""" & vbCrLf)
                            Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
                            Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
                            Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
                            Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
                            Response.Write("			    </td>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
                        Else
                            ' There are hidden components in the expression that ARE owned by the current user.
                            ' Need to make the expression hidden too.
                            If (Request.Form("validateAccess") <> "HD") Then
                                fDisplay = True
	
                                Response.Write("			  <tr>" & vbCrLf)
                                Response.Write("					<td width=20></td>" & vbCrLf)
                                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                                Response.Write("						<H3>Error Saving " & sUtilType & "</H3>" & vbCrLf)
                                Response.Write("			    </td>" & vbCrLf)
                                Response.Write("					<td width=20></td>" & vbCrLf)
                                Response.Write("			  </tr>" & vbCrLf)
                                Response.Write("			  <tr>" & vbCrLf)
                                Response.Write("					<td width=20></td>" & vbCrLf)
                                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                                Response.Write("						 The following calculations and filters have been made hidden. The " & sUtilType2 & " will now be made hidden." & vbCrLf)
                                Response.Write("					<td width=20></td>" & vbCrLf)
                                Response.Write("			  </tr>" & vbCrLf)
                                Response.Write("			  <tr>" & vbCrLf)
                                Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                                Response.Write("			  </tr>" & vbCrLf)
											
                                iIndex = InStr(sHiddenOwnerDescs, "	")
                                Do While iIndex > 0
                                    sDesc = Left(sHiddenOwnerDescs, iIndex - 1)

                                    Response.Write("			  <tr>" & vbCrLf)
                                    Response.Write("					<td width=20></td>" & vbCrLf)
                                    Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                                    Response.Write("						 " & sDesc & vbCrLf)
                                    Response.Write("			    </td>" & vbCrLf)
                                    Response.Write("					<td width=20></td>" & vbCrLf)
                                    Response.Write("			  </tr>" & vbCrLf)
												
                                    sHiddenOwnerDescs = Mid(sHiddenOwnerDescs, iIndex + 1)
                                    iIndex = InStr(sHiddenOwnerDescs, "	")
                                Loop

                                Response.Write("			  <tr>" & vbCrLf)
                                Response.Write("					<td width=20></td>" & vbCrLf)
                                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                                Response.Write("						 " & sHiddenOwnerDescs & vbCrLf)
                                Response.Write("			    </td>" & vbCrLf)
                                Response.Write("					<td width=20></td>" & vbCrLf)
                                Response.Write("			  </tr>" & vbCrLf)

                                Response.Write("			  <tr>" & vbCrLf)
                                Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                                Response.Write("			  </tr>" & vbCrLf)
                                Response.Write("			  <tr> " & vbCrLf)
                                Response.Write("					<td width=20></td>" & vbCrLf)
                                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                                Response.Write("    				    <INPUT TYPE=button VALUE=Close class=""btn"" NAME=makeHidden style=""WIDTH: 80px"" width=80 id=makeHidden" & vbCrLf)
                                Response.Write("    				        OnClick=""makeHidden()""" & vbCrLf)
                                Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
                                Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
                                Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
                                Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
                                Response.Write("			    </td>" & vbCrLf)
                                Response.Write("					<td width=20></td>" & vbCrLf)
                                Response.Write("			  </tr>" & vbCrLf)
                            End If
                        End If
                    Else
                        ' Current user is NOT the owner of the filter/calc.
						fDisplay = true

                        If (Len(sHiddenNotOwnerKeys) > 0) Then
                            ' There are hidden components in the expression that are NOT owned by the current user.
                            ' Cannot edit the expression as it too must now be hidden.
                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                            Response.Write("						<H3>Error Saving " & sUtilType & "</H3>" & vbCrLf)
                            Response.Write("			    </td>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                            Response.Write("						 The following calculations and filters have been made hidden by another user. Cannot make any modifications to this " & sUtilType2 & " definition." & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
										
                            iIndex = InStr(sHiddenNotOwnerDescs, "	")
                            Do While iIndex > 0
                                sDesc = Left(sHiddenNotOwnerDescs, iIndex - 1)

                                Response.Write("			  <tr>" & vbCrLf)
                                Response.Write("					<td width=20></td>" & vbCrLf)
                                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                                Response.Write("						 " & sDesc & vbCrLf)
                                Response.Write("			    </td>" & vbCrLf)
                                Response.Write("					<td width=20></td>" & vbCrLf)
                                Response.Write("			  </tr>" & vbCrLf)
											
                                sHiddenNotOwnerDescs = Mid(sHiddenNotOwnerDescs, iIndex + 1)
                                iIndex = InStr(sHiddenNotOwnerDescs, "	")
                            Loop

                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                            Response.Write("						 " & sHiddenNotOwnerDescs & vbCrLf)
                            Response.Write("			    </td>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)

                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
                            Response.Write("			  <tr> " & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                            Response.Write("    				    <INPUT TYPE=button VALUE=Close class=""btn"" NAME=ReturnToDefSel style=""WIDTH: 80px"" width=80 id=ReturnToDefSel" & vbCrLf)
                            Response.Write("    				        OnClick=""returnToDefSel()""" & vbCrLf)
                            Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
                            Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
                            Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
                            Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
                            Response.Write("			    </td>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
                        Else
                            ' There are hidden components in the expression that ARE owned by the current user.
                            ' Need to remove the hidden components.	
                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                            Response.Write("						<H3>Error Saving " & sUtilType & "</H3>" & vbCrLf)
                            Response.Write("			    </td>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                            Response.Write("						 The following calculations and filters have been made hidden, and will be removed from the " & sUtilType2 & " definition." & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
										
                            iIndex = InStr(sHiddenOwnerDescs, "	")
                            Do While iIndex > 0
                                sDesc = Left(sHiddenOwnerDescs, iIndex - 1)

                                Response.Write("			  <tr>" & vbCrLf)
                                Response.Write("					<td width=20></td>" & vbCrLf)
                                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                                Response.Write("						 " & sDesc & vbCrLf)
                                Response.Write("			    </td>" & vbCrLf)
                                Response.Write("					<td width=20></td>" & vbCrLf)
                                Response.Write("			  </tr>" & vbCrLf)
											
                                sHiddenOwnerDescs = Mid(sHiddenOwnerDescs, iIndex + 1)
                                iIndex = InStr(sHiddenOwnerDescs, "	")
                            Loop

                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                            Response.Write("						 " & sHiddenOwnerDescs & vbCrLf)
                            Response.Write("			    </td>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)

                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
                            Response.Write("			  <tr> " & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                            Response.Write("    				    <INPUT TYPE=button VALUE=Close class=""btn"" NAME=RemoveComponents style=""WIDTH: 80px"" width=80 id=RemoveComponents" & vbCrLf)
                            Response.Write("    				        OnClick=""removeComponents(3)""" & vbCrLf)
                            Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
                            Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
                            Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
                            Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
                            Response.Write("			    </td>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
                        End If
					end if					
				end if
			end if
		end if
	
		if (iErrorCode = 3) and _
			(fDisplay = false) then
			' 3 = Modified by another user (still writable). Overwrite ? 
			fDisplay = true

            Response.Write("			  <tr>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
            Response.Write("						<H3>Error Saving " & sUtilType & "</H3>" & vbCrLf)
            Response.Write("			    </td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			  </tr>" & vbCrLf)
            Response.Write("			  <tr>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
            Response.Write("						The " & sUtilType2 & " has been amended by another user. Would you like to overwrite this definition ?" & vbCrLf)
            Response.Write("			    </td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			  </tr>" & vbCrLf)
            Response.Write("			  <tr>" & vbCrLf)
            Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
            Response.Write("			  </tr>" & vbCrLf)
            Response.Write("			  <tr> " & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=right> " & vbCrLf)
            Response.Write("    				    <INPUT TYPE=button VALUE=Yes class=""btn"" NAME=btnYes style=""WIDTH: 80px"" width=80 id=btnYes" & vbCrLf)
            Response.Write("    				        OnClick=""overwrite()""" & vbCrLf)
            Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
            Response.Write("			    </td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=left> " & vbCrLf)
            Response.Write("    				    <INPUT TYPE=button VALUE=No class=""btn"" NAME=btnNo style=""WIDTH: 80px"" width=80 id=btnNo" & vbCrLf)
            Response.Write("    				        OnClick=""cancelClick()""" & vbCrLf)
            Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
            Response.Write("			    </td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("				</tr>" & vbCrLf)
        End If
	end if	
  
    
	if Request.form("validatePass") = 2 then
		' Get the server DLL to validate the expression definition
        objExpression = CreateObject("COAIntServer.Expression")

		' Pass required info to the DLL
		objExpression.Username = session("username")
        CallByName(objExpression, "Connection", CallType.Let, Session("databaseConnection"))
        
		if Request.form("validateUtilType") = 11 then
			iExprType = 11
			iReturnType = 3
		else
			iExprType = 10
			iReturnType = 0
		end if
				
        fOK = objExpression.Initialise(CLng(Request.Form("validateBaseTableID")), _
            CLng(Request.Form("validateUtilID")), CInt(iExprType), CInt(iReturnType))

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
                Response.Write("						<H3>Error Saving " & sUtilType & "</H3>" & vbCrLf)
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
                Response.Write("    				    <INPUT TYPE=button VALUE=Close class=""btn"" NAME=Cancel style=""WIDTH: 80px"" width=80 id=Cancel" & vbCrLf)
                Response.Write("    				        OnClick=""cancelClick()""" & vbCrLf)
                Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
                Response.Write("			    </td>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			  </tr>" & vbCrLf)
            Else
                iReturnType = objExpression.returnType
            End If
        End If
			
        objExpression = Nothing
	
		if not fDisplay then
			' Check if the expression return type has changed. 
			' If so, check if it can be.

		  If Request.form("validateUtilID") > 0 Then
                objExpression = CreateObject("COAIntServer.Expression")
                iOriginalReturnType = objExpression.ExistingExpressionReturnType(CLng(Request.Form("validateUtilID")))
                objExpression = Nothing
					
				If iReturnType <> iOriginalReturnType Then
                    cmdDefPropRecords = CreateObject("ADODB.Command")
					cmdDefPropRecords.CommandText = "sp_ASRIntDefUsage"
					cmdDefPropRecords.CommandType = 4 ' Stored Procedure

                    cmdDefPropRecords.ActiveConnection = Session("databaseConnection")

                    prmType = cmdDefPropRecords.CreateParameter("type", 3, 1) ' 3=integer, 1=input
                    cmdDefPropRecords.Parameters.Append(prmType)
					prmType.value = cleanNumeric(Request.form("validateUtilType"))

                    prmID = cmdDefPropRecords.CreateParameter("id", 3, 1) ' 3=integer, 1=input
                    cmdDefPropRecords.Parameters.Append(prmID)
					prmID.value = cleanNumeric(Request.form("validateUtilID"))

                    Err.Clear()
                    rsDefProp = cmdDefPropRecords.Execute

					if not (rsDefProp.BOF and rsDefProp.EOF) then
                        fDisplay = True
                        Response.Write("			  <tr>" & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                        Response.Write("						<H3>Error Saving " & sUtilType & "</H3>" & vbCrLf)
                        Response.Write("			    </td>" & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			  </tr>" & vbCrLf)
                        Response.Write("			  <tr>" & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                        Response.Write("						 The return type cannot be changed, as the " & sUtilType2 & " is currently being used in the following definitions." & vbCrLf)
                        Response.Write("			    </td>" & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			  </tr>" & vbCrLf)
                        Response.Write("			  <tr>" & vbCrLf)
                        Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                        Response.Write("			  </tr>" & vbCrLf)
						
						do while not rsDefProp.EOF
							sDescription = rsDefProp.Fields("description").Value
							sDescription = replace(sDescription, "<", "&lt;")
							sDescription = replace(sDescription, ">", "&gt;")

                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                            Response.Write("						 " & sDescription & vbCrLf)
                            Response.Write("			    </td>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)

                            rsDefProp.MoveNext()
                        Loop

                        Response.Write("			  <tr>" & vbCrLf)
                        Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                        Response.Write("			  </tr>" & vbCrLf)
                        Response.Write("			  <tr> " & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                        Response.Write("    				    <INPUT TYPE=button VALUE=Close class=""btn"" NAME=Cancel style=""WIDTH: 80px"" width=80 id=Cancel" & vbCrLf)
                        Response.Write("    				        OnClick=""cancelClick()""" & vbCrLf)
                        Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
                        Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
                        Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
                        Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
                        Response.Write("			    </td>" & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			  </tr>" & vbCrLf)
					end if
									   
                    rsDefProp = Nothing
                    cmdDefPropRecords = Nothing
			  End If
			End If		
		end if
	end if
		
	if Request.form("validatePass") = 3 then
	
		if (Request.form("validateUtilID") > 0) and _
			(ucase(Request.form("validateOwner")) = ucase(session("username"))) and _
			(Request.form("validateAccess") = "HD") and _
			(Request.form("validateOriginalAccess") <> "HD") then
			' Check if the expression can be made hidden.

            cmdCheckHidden = CreateObject("ADODB.Command")
			cmdCheckHidden.CommandText = "sp_ASRIntCheckCanMakeHidden"
			cmdCheckHidden.CommandType = 4 ' Stored Procedure
			cmdCheckHidden.CommandTimeout = 0

            cmdCheckHidden.ActiveConnection = Session("databaseConnection")

            prmType = cmdCheckHidden.CreateParameter("type", 3, 1)  ' 3=integer, 1=input
            cmdCheckHidden.Parameters.Append(prmType)
			prmType.value = cleanNumeric(Request.form("validateUtilType"))

            prmID = cmdCheckHidden.CreateParameter("id", 3, 1) ' 3=integer, 1=input
            cmdCheckHidden.Parameters.Append(prmID)
			prmID.value = cleanNumeric(Request.form("validateUtilID"))

            prmResult = cmdCheckHidden.CreateParameter("result", 3, 2) ' 3=integer, 2=output
            cmdCheckHidden.Parameters.Append(prmResult)

            prmMsg = cmdCheckHidden.CreateParameter("msg", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
            cmdCheckHidden.Parameters.Append(prmMsg)

            Err.Clear()
			cmdCheckHidden.Execute

			IF cmdCheckHidden.Parameters("result").Value = 1 then
				' calc/filter used only in utilities owned by the current user - we then need to prompt the user if they want to make these utilities hidden too.
				fDisplay = true
				sHiddenErrorMsg = "Making this " & sUtilType2 & " hidden will automatically make the following definition(s), of which you are the owner, hidden also :" & _
					"<BR><BR>" & _
					cmdCheckHidden.Parameters("msg").Value & _
					"<BR><BR>"  & _
					"Do you wish to continue ?"
				
                Response.Write("			  <tr>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                Response.Write("						<H3>Error Saving " & sUtilType & "</H3>" & vbCrLf)
                Response.Write("			    </td>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			  </tr>" & vbCrLf)
                Response.Write("			  <tr>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                Response.Write("						 " & sHiddenErrorMsg & vbCrLf)
                Response.Write("			    </td>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			  </tr>" & vbCrLf)
                Response.Write("			  <tr>" & vbCrLf)
                Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                Response.Write("			  </tr>" & vbCrLf)
                Response.Write("			  <tr> " & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			    <td align=right> " & vbCrLf)
                Response.Write("    				    <INPUT TYPE=button VALUE=Yes class=""btn"" NAME=btnYes style=""WIDTH: 80px"" width=80 id=btnYes" & vbCrLf)
                Response.Write("    				        OnClick=""overwrite()""" & vbCrLf)
                Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
                Response.Write("			    </td>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			    <td align=left> " & vbCrLf)
                Response.Write("    				    <INPUT TYPE=button VALUE=No class=""btn"" NAME=btnNo style=""WIDTH: 80px"" width=80 id=btnNo" & vbCrLf)
                Response.Write("    				        OnClick=""cancelClick()""" & vbCrLf)
                Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
                Response.Write("			    </td>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("				</tr>" & vbCrLf)
            End If

			IF (cmdCheckHidden.Parameters("result").Value = 2) or _
				(cmdCheckHidden.Parameters("result").Value = 3)  or _
				(cmdCheckHidden.Parameters("result").Value = 4) then
				' calc/filter used in utilities which are in batch jobs not owned by the current user - Cannot therefore make the calc/filter hidden.
				IF (cmdCheckHidden.Parameters("result").Value = 2) then 
					sHiddenErrorMsg = "This " & sUtilType2 & " cannot be made hidden as it is used in definition(s) which are included in the following batch jobs of which you are not the owner :" & _
						"<BR><BR>" & _
						cmdCheckHidden.Parameters("msg").Value
				else
					IF (cmdCheckHidden.Parameters("result").Value = 3) then 
						sHiddenErrorMsg = "This " & sUtilType2 & " cannot be made hidden as it is used in definition(s), of which you are not the owner :" & _
							"<BR><BR>" & _
							cmdCheckHidden.Parameters("msg").Value
					else
						sHiddenErrorMsg = "This " & sUtilType2 & " cannot be made hidden as it is used in definition(s) which are included in the following batch jobs which are scheduled to be run by other user groups :" & _
							"<BR><BR>" & _
							cmdCheckHidden.Parameters("msg").Value
					end if
				end if
				fDisplay = true
				
                Response.Write("			  <tr>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                Response.Write("						<H3>Error Saving " & sUtilType & "</H3>" & vbCrLf)
                Response.Write("			    </td>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			  </tr>" & vbCrLf)
                Response.Write("			  <tr>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                Response.Write("						 " & sHiddenErrorMsg & vbCrLf)
                Response.Write("			    </td>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			  </tr>" & vbCrLf)
                Response.Write("			  <tr>" & vbCrLf)
                Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                Response.Write("			  </tr>" & vbCrLf)
                Response.Write("			  <tr> " & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                Response.Write("    				    <INPUT TYPE=button VALUE=Close class=""btn"" NAME=Cancel style=""WIDTH: 80px"" width=80 id=Cancel" & vbCrLf)
                Response.Write("    				        OnClick=""cancelClick()""" & vbCrLf)
                Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
                Response.Write("			    </td>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			  </tr>" & vbCrLf)
            End If
        End If
    End If
	
    Response.Write("<INPUT type=hidden id=txtDisplay name=txtDisplay value=" & fDisplay & ">" & vbCrLf)
%>
			  <tr height=10> 
					<td colspan=5></td>
			  </tr>
			</table>
		</TD>
	</TR>
</table>
            </div>
            

<FORM id=frmValidate name=frmValidate method=post action=util_validate_expression style="visibility:hidden;display:none">
	<INPUT type=hidden id="validatePass" name=validatePass value=<%=Request.form("validatePass")%>>
	<INPUT type=hidden id="validateUtilID" name=validateUtilID value=<%=Request.form("validateUtilID")%>>
	<INPUT type=hidden id="validateUtilType" name=validateUtilType value=<%=Request.form("validateUtilType")%>>
	<INPUT type=hidden id="validateAccess" name=validateAccess value=<%=Request.form("validateAccess")%>>
	<INPUT type=hidden id="validateOriginalAccess" name=validateOriginalAccess value=<%=Request.form("validateOriginalAccess")%>>
	<INPUT type=hidden id="validateOwner" name=validateOwner value=<%=Request.form("validateOwner")%>>

	<INPUT type=hidden id=components1 name=components1 value="<%=Request.form("components1")%>">
	<INPUT type=hidden id=validateBaseTableID name=validateBaseTableID value=<%=Request.form("validateBaseTableID")%>>
</FORM>
    

</body>
</html>


<script type="text/javascript">
    util_validate_window_onload();
</script>
