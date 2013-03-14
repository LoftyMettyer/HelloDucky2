<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<div data-framesource="util_delete">

	<% 
		'Dim sSQL As String
		Dim sErrorDescription As String
		Dim sPrimaryTableName As String = ""
		Dim sCheckStatus As String = ""
	
		Dim sPrimaryIDColumnName As String = ""
		Dim sUtilTypeName As String = ""
	
		If Session("utiltype") = 1 Then
			' Cross Tabs
			sPrimaryTableName = "AsrSysCrossTab"
			sPrimaryIDColumnName = "CrossTabID"
			sUtilTypeName = "cross tab"

		ElseIf Session("utiltype") = 2 Then
			' Custom Reports
			sPrimaryTableName = "ASRSysCustomReportsName"
			sPrimaryIDColumnName = "id"
			sUtilTypeName = "report"

		ElseIf Session("utiltype") = 3 Then
			' Data Transfer
			sPrimaryTableName = "ASRSysDataTransferName"
			sPrimaryIDColumnName = "DataTransferID"
			sUtilTypeName = "data transfer"

		ElseIf Session("utiltype") = 4 Then
			' Export
			sPrimaryTableName = "ASRSysExportName"
			sPrimaryIDColumnName = "id"
			sUtilTypeName = "export"

		ElseIf Session("utiltype") = 5 Then
			' Global Add
			sPrimaryTableName = "ASRSysGlobalFunctions"
			sPrimaryIDColumnName = "FunctionID"
			sUtilTypeName = "global add"

		ElseIf Session("utiltype") = 6 Then
			' Global Update
			sPrimaryTableName = "ASRSysGlobalFunctions"
			sPrimaryIDColumnName = "FunctionID"
			sUtilTypeName = "global update"

		ElseIf Session("utiltype") = 7 Then
			' Global Delete
			sPrimaryTableName = "ASRSysGlobalFunctions"
			sPrimaryIDColumnName = "FunctionID"
			sUtilTypeName = "global delete"

		ElseIf Session("utiltype") = 8 Then
			' Import
			sPrimaryTableName = "ASRSysImportName"
			sPrimaryIDColumnName = "id"
			sUtilTypeName = "import"

		ElseIf Session("utiltype") = 9 Then
			' Mail Merge
			sPrimaryTableName = "AsrSysMailMergeName"
			sPrimaryIDColumnName = "MailMergeID"
			sUtilTypeName = "mail merge"

		ElseIf Session("utiltype") = 10 Then
			' Picklists
			sPrimaryTableName = "ASRSysPickListName"
			sPrimaryIDColumnName = "picklistID"
			sUtilTypeName = "picklist"

		ElseIf Session("utiltype") = 11 Then
			' Filters
			sPrimaryTableName = "ASRSysExpressions"
			sPrimaryIDColumnName = "exprID"
			sUtilTypeName = "filter"

		ElseIf Session("utiltype") = 12 Then
			' Calculations
			sPrimaryTableName = "ASRSysExpressions"
			sPrimaryIDColumnName = "exprID"
			sUtilTypeName = "calculation"

		ElseIf Session("utiltype") = 17 Then
			' Calendar Reports
			sPrimaryTableName = "ASRSysCalendarReports"
			sPrimaryIDColumnName = "id"
			sUtilTypeName = "calendar report"
	
		End If
	
		If Len(sPrimaryTableName) > 0 Then
			Dim cmdDeleteCheck = CreateObject("ADODB.Command")
			'Server.CreateObject("ADODB.Command")
			cmdDeleteCheck.CommandText = "spASRIntDeleteCheck"
			cmdDeleteCheck.CommandType = 4 ' Stored Procedure
			cmdDeleteCheck.ActiveConnection = Session("databaseConnection")

			Dim prmUtilType = cmdDeleteCheck.CreateParameter("utilType", 3, 1) ' 3=integer,1=input
			cmdDeleteCheck.Parameters.Append(prmUtilType)
			prmUtilType.value = CleanNumeric(Session("utiltype"))

			Dim prmUtilID As Object = cmdDeleteCheck.CreateParameter("utilID", 3, 1) ' 3=integer,1=input
			cmdDeleteCheck.Parameters.Append(prmUtilID)
			prmUtilID.value = CleanNumeric(Session("utilid"))

			Dim prmDeleted = cmdDeleteCheck.CreateParameter("deleted", 11, 2)	'11=bit, 2=output
			cmdDeleteCheck.Parameters.Append(prmDeleted)

			Dim prmAccess = cmdDeleteCheck.CreateParameter("access", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDeleteCheck.Parameters.Append(prmAccess)

			Err.Clear()
			cmdDeleteCheck.Execute()

			If cmdDeleteCheck.Parameters("deleted").Value = True Then
				sCheckStatus = CType(("'" & Session("utilname") & "' " & sUtilTypeName & " has been deleted by another user."), String)
			ElseIf cmdDeleteCheck.Parameters("access").Value = "HD" Then
				sCheckStatus = CType(("'" & Session("utilname") & "' " & sUtilTypeName & " has been made hidden by another user."), String)
			ElseIf cmdDeleteCheck.Parameters("access").Value = "RO" Then
				sCheckStatus = CType(("'" & Session("utilname") & "' " & sUtilTypeName & " has been made read only by another user."), String)
			End If
		
			cmdDeleteCheck = Nothing
			
			If Len(sCheckStatus) > 0 Then
				Session("confirmtext") = sCheckStatus
				Session("confirmtitle") = "OpenHR Intranet"
				Session("followpage") = "defsel"
				Response.Redirect("confirmok")
			End If
		
			
			' Check was okay, so go ahead and delete the utility.
			Dim cmdDelete = CreateObject("ADODB.Command")
			cmdDelete.CommandText = "sp_ASRIntDeleteUtility"
			cmdDelete.CommandType = 4	' Stored Procedure
			cmdDelete.ActiveConnection = Session("databaseConnection")
			
			prmUtilType = Nothing
			prmUtilType = cmdDelete.CreateParameter("utilType", 3, 1)	' 3=integer,1=input
			cmdDelete.Parameters.Append(prmUtilType)
			prmUtilType.value = CleanNumeric(Session("utiltype"))

			prmUtilID = Nothing
			prmUtilID = cmdDelete.CreateParameter("utilID", 3, 1)	' 3=integer,1=input
			cmdDelete.Parameters.Append(prmUtilID)
			prmUtilID.value = CleanNumeric(Session("utilid"))

			Err.Clear()
			cmdDelete.Execute()

			If Err.Number = 0 Then
				Session("confirmtext") = "'" & Session("utilname") & "' " & sUtilTypeName & " has been deleted."
				Session("confirmtitle") = "Delete Confirmation"
				Session("followpage") = "defsel"
				Response.Redirect("confirmok")
			End If
		
			sErrorDescription = Err.Description
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
			Response.Write("						<INPUT TYPE=button VALUE=OK NAME=GoBack OnClick=location.href=""defsel"" class=""btn"" style=""WIDTH: 80px"" width=80 id=cmdGoBack>" & vbCrLf)
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
		End If
	%>
</div>
