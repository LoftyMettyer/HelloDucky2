<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<div data-framesource="util_delete">

<% 
	dim sSQL 
	dim sErrorDescription
	dim sPrimaryTableName
	dim sCheckStatus
    Dim sPrimaryIDColumnName As String
    Dim sUtilTypeName As String
    Dim cmdDeleteCheck 
    Dim prmUtilType 
    Dim prmUtilID
    Dim prmDeleted
    Dim prmAccess
    Dim cmdDelete
    
    sErrorDescription = ""
	sPrimaryTableName = ""
	sCheckStatus = ""


	if session("utiltype") = 1 then
		' Cross Tabs
		sPrimaryTableName = "AsrSysCrossTab"
		sPrimaryIDColumnName = "CrossTabID"
		sUtilTypeName = "cross tab"

	elseif session("utiltype") = 2 then
		' Custom Reports
		sPrimaryTableName = "ASRSysCustomReportsName"
		sPrimaryIDColumnName = "id"
		sUtilTypeName = "report"

	elseif session("utiltype") = 3 then
		' Data Transfer
		sPrimaryTableName = "ASRSysDataTransferName"
		sPrimaryIDColumnName = "DataTransferID"
		sUtilTypeName = "data transfer"

	elseif session("utiltype") = 4 then
		' Export
		sPrimaryTableName = "ASRSysExportName"
		sPrimaryIDColumnName = "id"
		sUtilTypeName = "export"

	elseif session("utiltype") = 5 then
		' Global Add
		sPrimaryTableName = "ASRSysGlobalFunctions"
		sPrimaryIDColumnName = "FunctionID"
		sUtilTypeName = "global add"

	elseif session("utiltype") = 6 then
		' Global Update
		sPrimaryTableName = "ASRSysGlobalFunctions"
		sPrimaryIDColumnName = "FunctionID"
		sUtilTypeName = "global update"

	elseif session("utiltype") = 7 then
		' Global Delete
		sPrimaryTableName = "ASRSysGlobalFunctions"
		sPrimaryIDColumnName = "FunctionID"
		sUtilTypeName = "global delete"

	elseif session("utiltype") = 8 then
		' Import
		sPrimaryTableName = "ASRSysImportName"
		sPrimaryIDColumnName = "id"
		sUtilTypeName = "import"

	elseif session("utiltype") = 9 then
		' Mail Merge
		sPrimaryTableName = "AsrSysMailMergeName"
		sPrimaryIDColumnName = "MailMergeID"
		sUtilTypeName = "mail merge"

	elseif session("utiltype") = 10 then
		' Picklists
		sPrimaryTableName = "ASRSysPickListName"
		sPrimaryIDColumnName = "picklistID"
		sUtilTypeName = "picklist"

	elseif session("utiltype") = 11 then
		' Filters
		sPrimaryTableName = "ASRSysExpressions"
		sPrimaryIDColumnName = "exprID"
		sUtilTypeName = "filter"

	elseif session("utiltype") = 12 then
		' Calculations
		sPrimaryTableName = "ASRSysExpressions"
		sPrimaryIDColumnName = "exprID"
		sUtilTypeName = "calculation"

	elseif session("utiltype") = 17 then
		' Calendar Reports
		sPrimaryTableName = "ASRSysCalendarReports"
		sPrimaryIDColumnName = "id"
		sUtilTypeName = "calendar report"
	
	end if
	
	if len(sPrimaryTableName) > 0 then
        cmdDeleteCheck = Server.CreateObject("ADODB.Command")
		cmdDeleteCheck.CommandText = "spASRIntDeleteCheck"
		cmdDeleteCheck.CommandType = 4 ' Stored Procedure
        cmdDeleteCheck.ActiveConnection = Session("databaseConnection")

        prmUtilType = cmdDeleteCheck.CreateParameter("utilType", 3, 1) ' 3=integer,1=input
        cmdDeleteCheck.Parameters.Append(prmUtilType)
		prmUtilType.value = cleanNumeric(session("utiltype"))

        prmUtilID = cmdDeleteCheck.CreateParameter("utilID", 3, 1) ' 3=integer,1=input
        cmdDeleteCheck.Parameters.Append(prmUtilID)
		prmUtilID.value = cleanNumeric(session("utilid"))

        prmDeleted = cmdDeleteCheck.CreateParameter("deleted", 11, 2) '11=bit, 2=output
        cmdDeleteCheck.Parameters.Append(prmDeleted)

        prmAccess = cmdDeleteCheck.CreateParameter("access", 200, 2, 8000) '200=varchar, 2=output, 8000=size
        cmdDeleteCheck.Parameters.Append(prmAccess)

        Err.Clear()
		cmdDeleteCheck.Execute

		if cmdDeleteCheck.Parameters("deleted").Value = true then
			sCheckStatus = "'" & session("utilname") & "' " & sUtilTypeName & " has been deleted by another user."
		elseif cmdDeleteCheck.Parameters("access").Value = "HD" then
			sCheckStatus = "'" & session("utilname") & "' " & sUtilTypeName & " has been made hidden by another user."
		elseif cmdDeleteCheck.Parameters("access").Value = "RO" then
			sCheckStatus = "'" & session("utilname") & "' " & sUtilTypeName & " has been made read only by another user."
		end if	
		
        cmdDeleteCheck = Nothing
		
		if len(sCheckStatus) > 0 then
			session("confirmtext") = sCheckStatus
			session("confirmtitle") = "OpenHR Intranet"
            Session("followpage") = "defsel"
            Response.Redirect("confirmok")
		end if
		
		' Check was okay, so go ahead and delete the utility.
        cmdDelete = Server.CreateObject("ADODB.Command")
		cmdDelete.CommandText = "sp_ASRIntDeleteUtility"
		cmdDelete.CommandType = 4 ' Stored Procedure
        cmdDelete.ActiveConnection = Session("databaseConnection")

        prmUtilType = cmdDelete.CreateParameter("utilType", 3, 1) ' 3=integer,1=input
        cmdDelete.Parameters.Append(prmUtilType)
		prmUtilType.value = cleanNumeric(session("utiltype"))

        prmUtilID = cmdDelete.CreateParameter("utilID", 3, 1) ' 3=integer,1=input
        cmdDelete.Parameters.Append(prmUtilID)
		prmUtilID.value = cleanNumeric(session("utilid"))

        Err.Clear()
		cmdDelete.Execute

        If Err.Number = 0 Then
            Session("confirmtext") = "'" & Session("utilname") & "' " & sUtilTypeName & " has been deleted."
            Session("confirmtitle") = "Delete Confirmation"
            Session("followpage") = "defsel"
            Response.Redirect("confirmok")
        End If
		
		sErrorDescription = err.description
	
		' If we are here, something has gone wrong, 
		' So display the header/stylesheet code, then tell the
        ' user what the problem was.

        Response.Write("<HTML>" & vbCrLf)
        Response.Write("	<HEAD>" & vbCrLf)
        Response.Write("		<META NAME=""GENERATOR"" Content=""Microsoft Visual Studio 6.0"">" & vbCrLf)
        Response.Write("		<LINK href=""OpenHR.css"" rel=stylesheet type=text/css >" & vbCrLf)
        Response.Write("		<TITLE>OpenHR Intranet</TITLE>" & vbCrLf)
        Response.Write("  <!--#INCLUDE FILE=""include/ctl_SetStyles.txt"" -->" & vbCrLf)
        Response.Write("	</HEAD>" & vbCrLf)
        Response.Write("	<BODY " & Session("BodyTag") & ">" & vbCrLf)
        Response.Write("	<table align=center class=""outline"" cellPadding=5 cellSpacing=0>" & vbCrLf)
        Response.Write("		<TR>" & vbCrLf)
        Response.Write("			<TD>" & vbCrLf)
        Response.Write("				<table class=""invisible"" cellspacing=0 cellpadding=0>" & vbCrLf)
        Response.Write("				  <tr> " & vbCrLf)
        Response.Write("				    <td colspan=3 height=10></td>" & vbCrLf)
        Response.Write("				  </tr>" & vbCrLf)
        Response.Write("				  <tr> " & vbCrLf)
        Response.Write("				    <td colspan=3 align=center> " & vbCrLf)
        Response.Write("							<H3>Error</H3>" & vbCrLf)
        Response.Write("				    </td>" & vbCrLf)
        Response.Write("				  </tr>" & vbCrLf)
        Response.Write("				  <tr> " & vbCrLf)
        Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
        Response.Write("				    <td> " & vbCrLf)
        Response.Write("							<H4>An error has occured whilst performing the delete operation.</H4>" & vbCrLf)
        Response.Write("				    </td>" & vbCrLf)
        Response.Write("				    <td width=20></td> " & vbCrLf)
        Response.Write("				  </tr>" & vbCrLf)
        Response.Write("				  <tr> " & vbCrLf)
        Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
        Response.Write("				    <td> " & vbCrLf)
        Response.Write(sErrorDescription & vbCrLf)
        Response.Write("			    </td>" & vbCrLf)
        Response.Write("			    <td width=20></td> " & vbCrLf)
        Response.Write("			  </tr>" & vbCrLf)
        Response.Write("			  <tr> " & vbCrLf)
        Response.Write("			    <td colspan=3 height=20></td>" & vbCrLf)
        Response.Write("			  </tr>" & vbCrLf)
        Response.Write("			  <tr> " & vbCrLf)
        Response.Write("			    <td colspan=3 height=10 align=center>" & vbCrLf)
        Response.Write("						<INPUT TYPE=button VALUE=OK NAME=GoBack OnClick=location.href=""defsel.asp"" class=""btn"" style=""WIDTH: 80px"" width=80 id=cmdGoBack>" & vbCrLf)
        Response.Write("                      onclick=""window.parent.parent.self.close();""" & vbCrLf)
        Response.Write("                      onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
        Response.Write("                      onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
        Response.Write("                      onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
        Response.Write("                      onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
        Response.Write("			    </td>" & vbCrLf)
        Response.Write("			  </tr>" & vbCrLf)
        Response.Write("			  <tr>" & vbCrLf)
        Response.Write("			    <td colspan=3 height=10></td>" & vbCrLf)
        Response.Write("			  </tr>" & vbCrLf)
        Response.Write("			</table>" & vbCrLf)
        Response.Write("    </td>" & vbCrLf)
        Response.Write("  </tr>" & vbCrLf)
        Response.Write("</table>" & vbCrLf)
        Response.Write("</BODY>" & vbCrLf)
        Response.Write("</HTML>" & vbCrLf)
        Response.End()
    Else
        Response.Redirect("default")
	end if
%>

</div>