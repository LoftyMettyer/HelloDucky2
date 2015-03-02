<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<div data-framesource="util_delete">

	<% 
		Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

		Dim sPrimaryTableName As String = ""
		Dim sCheckStatus As String = ""
	
		Dim sUtilTypeName As String = ""
	
		If Session("utiltype") = 1 Then
			' Cross Tabs
			sPrimaryTableName = "AsrSysCrossTab"
			sUtilTypeName = "cross tab"

		ElseIf Session("utiltype") = 2 Then
			' Custom Reports
			sPrimaryTableName = "ASRSysCustomReportsName"
			sUtilTypeName = "report"

		ElseIf Session("utiltype") = 3 Then
			' Data Transfer
			sPrimaryTableName = "ASRSysDataTransferName"
			sUtilTypeName = "data transfer"

		ElseIf Session("utiltype") = 4 Then
			' Export
			sPrimaryTableName = "ASRSysExportName"
			sUtilTypeName = "export"

		ElseIf Session("utiltype") = 5 Then
			' Global Add
			sPrimaryTableName = "ASRSysGlobalFunctions"
			sUtilTypeName = "global add"

		ElseIf Session("utiltype") = 6 Then
			' Global Update
			sPrimaryTableName = "ASRSysGlobalFunctions"
			sUtilTypeName = "global update"

		ElseIf Session("utiltype") = 7 Then
			' Global Delete
			sPrimaryTableName = "ASRSysGlobalFunctions"
			sUtilTypeName = "global delete"

		ElseIf Session("utiltype") = 8 Then
			' Import
			sPrimaryTableName = "ASRSysImportName"
			sUtilTypeName = "import"

		ElseIf Session("utiltype") = 9 Then
			' Mail Merge
			sPrimaryTableName = "AsrSysMailMergeName"
			sUtilTypeName = "mail merge"

		ElseIf Session("utiltype") = 10 Then
			' Picklists
			sPrimaryTableName = "ASRSysPickListName"
			sUtilTypeName = "picklist"

		ElseIf Session("utiltype") = 11 Then
			' Filters
			sPrimaryTableName = "ASRSysExpressions"
			sUtilTypeName = "filter"

		ElseIf Session("utiltype") = 12 Then
			' Calculations
			sPrimaryTableName = "ASRSysExpressions"
			sUtilTypeName = "calculation"

		ElseIf Session("utiltype") = 17 Then
			' Calendar Reports
			sPrimaryTableName = "ASRSysCalendarReports"
			sUtilTypeName = "calendar report"

		ElseIf Session("utiltype") = 35 Then
			' Cross Tabs
			sPrimaryTableName = "AsrSysCrossTab"
			sUtilTypeName = "9-box grid report"

		End If
	
		If Len(sPrimaryTableName) > 0 Then

			Try
				
				Dim prmDeleted = New SqlParameter("pfDeleted", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmAccess = New SqlParameter("psAccess", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
					
				objDataAccess.ExecuteSP("spASRIntDeleteCheck" _
							, New SqlParameter("piUtilityType", SqlDbType.Int) With {.Value = CleanNumeric(Session("utiltype"))} _
							, New SqlParameter("plngID", SqlDbType.Int) With {.Value = CleanNumeric(Session("utilid"))} _
							, prmDeleted _
							, prmAccess)
				
				If CBool(prmDeleted.Value) = True Then
					sCheckStatus = CType(("'" & Session("utilname") & "' " & sUtilTypeName & " has been deleted by another user."), String)
				ElseIf prmAccess.Value.ToString() = "HD" Then
					sCheckStatus = CType(("'" & Session("utilname") & "' " & sUtilTypeName & " has been made hidden by another user."), String)
				ElseIf prmAccess.Value.ToString() = "RO" Then
					sCheckStatus = CType(("'" & Session("utilname") & "' " & sUtilTypeName & " has been made read only by another user."), String)
				End If
			
				If Len(sCheckStatus) > 0 Then
					Session("confirmtext") = sCheckStatus
					Session("confirmtitle") = "OpenHR"
					Session("followpage") = "defsel"
				Else
				
					objDataAccess.ExecuteSP("sp_ASRIntDeleteUtility" _
											, New SqlParameter("piUtilType", SqlDbType.Int) With {.Value = CleanNumeric(Session("utiltype"))} _
											, New SqlParameter("piUtilID", SqlDbType.Int) With {.Value = CleanNumeric(Session("utilid"))})

					Session("confirmtext") = "'" & Session("utilname") & "' " & sUtilTypeName & " has been deleted."
					Session("confirmtitle") = "Delete Confirmation"
					Session("followpage") = "defsel"
					
					'Reset the utilid to 0
					Session("utilid") = 0
				End If

				Response.Redirect("confirmok")

			Catch ex As Exception

				Session("ErrorTitle") = "Login Page"
				Session("ErrorText") = "An error has occured whilst performing the delete operation:" & vbCrLf & ex.Message & "<p>" & vbCrLf
				Response.Redirect("FormError")
				
			End Try

		End If
	%>
</div>
